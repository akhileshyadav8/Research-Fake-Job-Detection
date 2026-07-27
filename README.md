# Fake Job Posting Detection - Research Codebase

This repository contains the source code, training pipeline, evaluation scripts, and paper deliverables for detecting Online Recruitment Fraud (ORF) using a hybrid **BERT-BiLSTM** model compared against various machine learning and deep learning baselines.

---

## 1. Prerequisites & Environment Setup

Make sure you have Python 3.8+ installed (Anaconda or Miniconda is highly recommended). 

Open your terminal or command prompt and install the required dependencies:

```bash
pip install torch transformers scikit-learn pandas numpy matplotlib seaborn joblib python-docx
```

*Note: The code will automatically detect if a GPU (CUDA) is available. If a GPU is present, it will train on CUDA; otherwise, it will automatically run on the CPU (with downsampling enabled to ensure it runs in under 5 minutes).*

---

## 2. Directory Structure

Ensure the directory structure matches the following layout:

```text
Research - Fake Job Detection/
│
├── data/
│   └── fake_job_postings.csv     # The Kaggle dataset file (already downloaded)
│
├── src/
│   ├── data_loader.py            # Preprocessing and text concatenation
│   ├── models.py                 # Neural network architectures
│   ├── baselines.py              # ML baseline algorithms (LR & RF)
│   ├── train.py                  # Training loops and PyTorch datasets
│   └── evaluate.py               # Coordination script for all models
│
├── paper/
│   ├── paper.tex                 # LaTeX paper source
│   ├── references.bib            # Citation database
│   ├── paper.docx                # Compiled Word Document
│   └── architecture_placeholder.jpg # Model flowchart diagram
│
├── results/                      # Directory where trained models and plots are saved
└── README.md                     # Running guide
```

---

## 3. How to Run the Code

1. Open your terminal/command prompt.
2. Navigate to the `src` directory:
   ```bash
   cd "d:\M.Sc (Data Science)\Research - Fake Job Detection\src"
   ```
3. Run the evaluation coordinator:
   ```bash
   python evaluate.py
   ```

---

## 4. Expected Outputs

Once `evaluate.py` completes, it will save the following assets to the `results/` folder:
- **`summary.csv`**: A table comparing the Accuracy, Precision, Recall, F1, and AUC metrics.
- **`roc_curves.png`**: ROC curves for all 5 models.
- **`performance_comparison.png`**: Performance metrics bar chart.
- **`confusion_matrix_proposed.png`**: Confusion matrix for the proposed BERT-BiLSTM framework.
- **`confusion_matrix_lstm.png`**: Confusion matrix for the standard Bi-LSTM baseline.
- **Model Checkpoints**: Trained weights saved as `lstm_best.pt`, `bert_standalone_best.pt`, and `bert_bilstm_best.pt`.
