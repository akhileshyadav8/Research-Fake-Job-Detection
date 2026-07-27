import os
import re
import collections
import numpy as np
import torch
import torch.nn as nn
from torch.utils.data import Dataset, DataLoader
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, roc_auc_score
from transformers import AutoTokenizer

# Build a simple vocabulary for standard LSTM
class Vocabulary:
    def __init__(self, max_vocab_size=10000):
        self.max_vocab_size = max_vocab_size
        self.word2idx = {"<pad>": 0, "<unk>": 1}
        self.idx2word = {0: "<pad>", 1: "<unk>"}
        
    def fit(self, texts):
        word_counts = collections.Counter()
        for text in texts:
            tokens = self._tokenize(text)
            word_counts.update(tokens)
            
        most_common = word_counts.most_common(self.max_vocab_size - 2)
        for idx, (word, _) in enumerate(most_common, start=2):
            self.word2idx[word] = idx
            self.idx2word[idx] = word
            
    def _tokenize(self, text):
        # Convert to lowercase and get words
        return re.findall(r'\w+', text.lower())
        
    def encode(self, text, max_len=256):
        tokens = self._tokenize(text)
        indexed = [self.word2idx.get(w, 1) for w in tokens]
        # Pad or truncate
        if len(indexed) < max_len:
            indexed += [0] * (max_len - len(indexed))
        else:
            indexed = indexed[:max_len]
        return torch.tensor(indexed, dtype=torch.long)
        
    def __len__(self):
        return len(self.word2idx)


# PyTorch Dataset for LSTM
class LSTMDataset(Dataset):
    def __init__(self, df, vocab, max_len=256):
        self.labels = df['fraudulent'].values
        texts = df['combined_text'].values
        print(f"Encoding {len(texts)} texts for LSTM...")
        self.encoded_texts = [vocab.encode(t, max_len=max_len) for t in texts]
        print("Encoding complete.")
        
    def __len__(self):
        return len(self.labels)
        
    def __getitem__(self, idx):
        return self.encoded_texts[idx], torch.tensor(self.labels[idx], dtype=torch.float)


# PyTorch Dataset for BERT/DistilBERT
class BERTDataset(Dataset):
    def __init__(self, df, tokenizer, max_len=256):
        self.labels = df['fraudulent'].values
        texts = df['combined_text'].tolist()
        print(f"Tokenizing {len(texts)} texts in batch...")
        self.encodings = tokenizer(
            texts,
            add_special_tokens=True,
            max_length=max_len,
            padding='max_length',
            truncation=True,
            return_attention_mask=True,
            return_tensors='pt'
        )
        print("Tokenization complete.")
        
    def __len__(self):
        return len(self.labels)
        
    def __getitem__(self, idx):
        return {
            'input_ids': self.encodings['input_ids'][idx],
            'attention_mask': self.encodings['attention_mask'][idx],
            'label': torch.tensor(self.labels[idx], dtype=torch.float)
        }


# Generic Trainer class
class Trainer:
    def __init__(self, model, device, lr=1e-3, pos_weight=None):
        self.model = model.to(device)
        self.device = device
        
        # Optimizer
        self.optimizer = torch.optim.Adam(model.parameters(), lr=lr)
        
        # Loss with pos_weight for imbalanced classes
        if pos_weight is not None:
            self.criterion = nn.BCEWithLogitsLoss(pos_weight=torch.tensor([pos_weight], device=device))
        else:
            self.criterion = nn.BCEWithLogitsLoss()
            
    def train_epoch(self, dataloader, is_bert=True):
        self.model.train()
        total_loss = 0
        
        for batch in dataloader:
            self.optimizer.zero_grad()
            
            if is_bert:
                input_ids = batch['input_ids'].to(self.device)
                attention_mask = batch['attention_mask'].to(self.device)
                labels = batch['label'].to(self.device)
                logits = self.model(input_ids, attention_mask)
            else:
                texts, labels = batch
                texts = texts.to(self.device)
                labels = labels.to(self.device)
                logits = self.model(texts)
                
            loss = self.criterion(logits, labels)
            loss.backward()
            self.optimizer.step()
            
            total_loss += loss.item()
            
        return total_loss / len(dataloader)
        
    def evaluate(self, dataloader, is_bert=True):
        self.model.eval()
        all_preds = []
        all_probs = []
        all_labels = []
        
        with torch.no_grad():
            for batch in dataloader:
                if is_bert:
                    input_ids = batch['input_ids'].to(self.device)
                    attention_mask = batch['attention_mask'].to(self.device)
                    labels = batch['label'].to(self.device)
                    logits = self.model(input_ids, attention_mask)
                else:
                    texts, labels = batch
                    texts = texts.to(self.device)
                    labels = labels.to(self.device)
                    logits = self.model(texts)
                    
                probs = torch.sigmoid(logits).cpu().numpy()
                preds = (probs >= 0.5).astype(int)
                
                all_probs.extend(probs)
                all_preds.extend(preds)
                all_labels.extend(labels.cpu().numpy())
                
        acc = accuracy_score(all_labels, all_preds)
        prec = precision_score(all_labels, all_preds, zero_division=0)
        rec = recall_score(all_labels, all_preds)
        f1 = f1_score(all_labels, all_preds)
        try:
            auc = roc_auc_score(all_labels, all_probs)
        except ValueError:
            auc = 0.5 # if only one class in batch/labels
            
        return acc, prec, rec, f1, auc, np.array(all_probs), np.array(all_preds)


def train_deep_models(train_df, val_df, test_df, model_type="lstm", max_len=256, epochs=5, batch_size=32, lr=1e-3, results_dir="results"):
    os.makedirs(results_dir, exist_ok=True)
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    print(f"Using device: {device} for training {model_type}")
    
    # Calculate pos_weight (neg_count / pos_count) for imbalance
    neg_count = sum(train_df['fraudulent'] == 0)
    pos_count = sum(train_df['fraudulent'] == 1)
    pos_weight = neg_count / max(1, pos_count)
    print(f"Class ratio in training data: {neg_count} Negative / {pos_count} Positive. Pos weight: {pos_weight:.2f}")
    
    from models import LSTMClassifier, BERT_BiLSTM, BERTStandalone
    
    if model_type == "lstm":
        # Fit vocabulary
        print("Fitting vocabulary for LSTM...")
        vocab = Vocabulary(max_vocab_size=10000)
        vocab.fit(train_df['combined_text'])
        
        train_dataset = LSTMDataset(train_df, vocab, max_len=max_len)
        val_dataset = LSTMDataset(val_df, vocab, max_len=max_len)
        test_dataset = LSTMDataset(test_df, vocab, max_len=max_len)
        
        train_loader = DataLoader(train_dataset, batch_size=batch_size, shuffle=True)
        val_loader = DataLoader(val_dataset, batch_size=batch_size, shuffle=False)
        test_loader = DataLoader(test_dataset, batch_size=batch_size, shuffle=False)
        
        model = LSTMClassifier(vocab_size=len(vocab))
        trainer = Trainer(model, device, lr=lr, pos_weight=pos_weight)
        is_bert = False
        
    elif model_type in ["bert_standalone", "bert_bilstm", "roberta_standalone"]:
        model_name = "roberta-base" if model_type == "roberta_standalone" else "distilbert-base-uncased"
        tokenizer = AutoTokenizer.from_pretrained(model_name)
        
        train_dataset = BERTDataset(train_df, tokenizer, max_len=max_len)
        val_dataset = BERTDataset(val_df, tokenizer, max_len=max_len)
        test_dataset = BERTDataset(test_df, tokenizer, max_len=max_len)
        
        train_loader = DataLoader(train_dataset, batch_size=batch_size, shuffle=True)
        val_loader = DataLoader(val_dataset, batch_size=batch_size, shuffle=False)
        test_loader = DataLoader(test_dataset, batch_size=batch_size, shuffle=False)
        
        if model_type == "bert_standalone":
            model = BERTStandalone(model_name="distilbert-base-uncased", freeze_bert=True)
            trainer = Trainer(model, device, lr=1e-3, pos_weight=pos_weight)
        elif model_type == "roberta_standalone":
            model = BERTStandalone(model_name="roberta-base", freeze_bert=True)
            trainer = Trainer(model, device, lr=1e-3, pos_weight=pos_weight)
        else: # bert_bilstm proposed model
            model = BERT_BiLSTM(model_name="distilbert-base-uncased", freeze_bert=True)
            trainer = Trainer(model, device, lr=1e-3, pos_weight=pos_weight)
            
        is_bert = True
        
    else:
        raise ValueError(f"Unknown model type: {model_type}")
        
    best_val_f1 = 0
    checkpoint_path = os.path.join(results_dir, f"{model_type}_best.pt")
    
    print(f"Starting training for {model_type}...")
    for epoch in range(epochs):
        train_loss = trainer.train_epoch(train_loader, is_bert=is_bert)
        val_acc, val_prec, val_rec, val_f1, val_auc, _, _ = trainer.evaluate(val_loader, is_bert=is_bert)
        
        print(f"Epoch {epoch+1}/{epochs} - Loss: {train_loss:.4f} | Val F1: {val_f1:.4f} (Rec: {val_rec:.4f}, Prec: {val_prec:.4f}) | Val Acc: {val_acc:.4f}")
        
        if val_f1 > best_val_f1:
            best_val_f1 = val_f1
            torch.save(model.state_dict(), checkpoint_path)
            print(f"  Saved new best model checkpoint to {checkpoint_path}")
            
    # Load best checkpoint and evaluate on test set
    print("Evaluating best model on test set...")
    model.load_state_dict(torch.load(checkpoint_path))
    test_acc, test_prec, test_rec, test_f1, test_auc, test_probs, test_preds = trainer.evaluate(test_loader, is_bert=is_bert)
    
    print(f"Test Results for {model_type} - Acc: {test_acc:.4f}, Precision: {test_prec:.4f}, Recall: {test_rec:.4f}, F1: {test_f1:.4f}, AUC: {test_auc:.4f}")
    
    # Save predictions
    results = {
        "accuracy": test_acc,
        "precision": test_prec,
        "recall": test_rec,
        "f1": test_f1,
        "auc": test_auc,
        "probs": test_probs,
        "preds": test_preds,
        "y_true": test_df['fraudulent'].values
    }
    
    return results

if __name__ == "__main__":
    from data_loader import get_splits
    csv_path = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\data\fake_job_postings.csv"
    train, val, test = get_splits(csv_path)
    # Quick sanity check with small epoch and small slice
    train_slice = train.sample(100, random_state=42)
    val_slice = val.sample(50, random_state=42)
    test_slice = test.sample(50, random_state=42)
    train_deep_models(train_slice, val_slice, test_slice, model_type="lstm", epochs=1)
