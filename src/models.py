import torch
import torch.nn as nn
from transformers import AutoModel

class LSTMClassifier(nn.Module):
    def __init__(self, vocab_size, embedding_dim=100, hidden_dim=128, output_dim=1, num_layers=1, dropout=0.3):
        super().__init__()
        self.embedding = nn.Embedding(vocab_size, embedding_dim, padding_idx=0)
        self.lstm = nn.LSTM(
            embedding_dim, 
            hidden_dim, 
            num_layers=num_layers, 
            bidirectional=True, 
            batch_first=True, 
            dropout=dropout if num_layers > 1 else 0.0
        )
        self.fc = nn.Linear(hidden_dim * 2, output_dim)
        self.dropout = nn.Dropout(dropout)
        
    def forward(self, text, text_lengths=None):
        # text: (batch_size, seq_len)
        embedded = self.dropout(self.embedding(text))
        
        # lstm_out: (batch_size, seq_len, hidden_dim * 2)
        lstm_out, (hidden, cell) = self.lstm(embedded)
        
        # Global Max Pooling over sequence dimension
        # pool_out: (batch_size, hidden_dim * 2)
        pool_out, _ = torch.max(lstm_out, dim=1)
        
        # Classify
        logits = self.fc(self.dropout(pool_out))
        return logits.squeeze(1)


class BERT_BiLSTM(nn.Module):
    def __init__(self, model_name="distilbert-base-uncased", lstm_hidden_dim=128, output_dim=1, num_layers=1, dropout=0.3, freeze_bert=True):
        super().__init__()
        # Load pre-trained transformer model
        self.transformer = AutoModel.from_pretrained(model_name)
        transformer_hidden_dim = self.transformer.config.hidden_size
        
        # Option to freeze transformer weights for faster training and less VRAM consumption
        if freeze_bert:
            for param in self.transformer.parameters():
                param.requires_grad = False
                
        self.lstm = nn.LSTM(
            transformer_hidden_dim,
            lstm_hidden_dim,
            num_layers=num_layers,
            bidirectional=True,
            batch_first=True,
            dropout=dropout if num_layers > 1 else 0.0
        )
        
        self.fc = nn.Sequential(
            nn.Dropout(dropout),
            nn.Linear(lstm_hidden_dim * 2, 64),
            nn.ReLU(),
            nn.Dropout(dropout),
            nn.Linear(64, output_dim)
        )
        self.dropout = nn.Dropout(dropout)
        
    def forward(self, input_ids, attention_mask):
        # input_ids: (batch_size, seq_len)
        # attention_mask: (batch_size, seq_len)
        
        # Get token embeddings from BERT/DistilBERT
        with torch.set_grad_enabled(self.transformer.training and any(p.requires_grad for p in self.transformer.parameters())):
            outputs = self.transformer(input_ids=input_ids, attention_mask=attention_mask)
            # last_hidden_state shape: (batch_size, seq_len, transformer_hidden_dim)
            last_hidden_state = outputs.last_hidden_state
            
        # Feed sequence of token embeddings into Bi-LSTM
        lstm_out, _ = self.lstm(last_hidden_state)
        
        # Global Max Pooling over the sequence dimension to aggregate sequence representation
        pooled, _ = torch.max(lstm_out, dim=1)
        
        # Classification
        logits = self.fc(pooled)
        return logits.squeeze(1)


class BERTStandalone(nn.Module):
    def __init__(self, model_name="distilbert-base-uncased", output_dim=1, dropout=0.3, freeze_bert=False):
        super().__init__()
        self.transformer = AutoModel.from_pretrained(model_name)
        transformer_hidden_dim = self.transformer.config.hidden_size
        
        if freeze_bert:
            for param in self.transformer.parameters():
                param.requires_grad = False
                
        self.fc = nn.Linear(transformer_hidden_dim, output_dim)
        self.dropout = nn.Dropout(dropout)
        
    def forward(self, input_ids, attention_mask):
        with torch.set_grad_enabled(self.transformer.training and any(p.requires_grad for p in self.transformer.parameters())):
            outputs = self.transformer(input_ids=input_ids, attention_mask=attention_mask)
            # Standalone DistilBERT uses the first token [CLS] representation (index 0) for classification
            cls_repr = outputs.last_hidden_state[:, 0, :]
            
        logits = self.fc(self.dropout(cls_repr))
        return logits.squeeze(1)
