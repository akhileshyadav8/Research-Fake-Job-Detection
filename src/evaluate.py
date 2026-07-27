import os
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
import pickle
import joblib
from sklearn.metrics import confusion_matrix, roc_curve, auc

def plot_roc_curves(all_results, save_path):
    plt.figure(figsize=(10, 8))
    sns.set_theme(style="whitegrid")
    
    for name, res in all_results.items():
        fpr, tpr, _ = roc_curve(res["y_true"], res["probs"])
        roc_auc = auc(fpr, tpr)
        plt.plot(fpr, tpr, lw=2, label=f'{name} (AUC = {roc_auc:.4f})')
        
    plt.plot([0, 1], [0, 1], color='navy', lw=2, linestyle='--')
    plt.xlim([0.0, 1.0])
    plt.ylim([0.0, 1.05])
    plt.xlabel('False Positive Rate', fontsize=12)
    plt.ylabel('True Positive Rate', fontsize=12)
    plt.title('Receiver Operating Characteristic (ROC) Curves', fontsize=14, fontweight='bold')
    plt.legend(loc="lower right", fontsize=11)
    plt.tight_layout()
    plt.savefig(save_path, dpi=300)
    plt.close()
    print(f"ROC curves saved to {save_path}")

def plot_confusion_matrix(y_true, y_pred, save_path, title):
    cm = confusion_matrix(y_true, y_pred)
    plt.figure(figsize=(6, 5))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', cbar=False,
                xticklabels=['Genuine', 'Fraudulent'],
                yticklabels=['Genuine', 'Fraudulent'],
                annot_kws={"size": 14})
    plt.ylabel('Actual Class', fontsize=12, fontweight='bold')
    plt.xlabel('Predicted Class', fontsize=12, fontweight='bold')
    plt.title(title, fontsize=13, fontweight='bold')
    plt.tight_layout()
    plt.savefig(save_path, dpi=300)
    plt.close()
    print(f"Confusion matrix saved to {save_path}")

def plot_comparison_bar_chart(all_results, save_path):
    metrics = ["Accuracy", "Precision", "Recall", "F1-score", "ROC-AUC"]
    data = []
    
    for name, res in all_results.items():
        data.append({
            "Model": name,
            "Accuracy": res["accuracy"],
            "Precision": res["precision"],
            "Recall": res["recall"],
            "F1-score": res["f1"],
            "ROC-AUC": res["auc"]
        })
        
    df_metrics = pd.DataFrame(data)
    df_melted = df_metrics.melt(id_vars="Model", var_name="Metric", value_name="Score")
    
    plt.figure(figsize=(12, 7))
    sns.set_theme(style="whitegrid")
    ax = sns.barplot(x="Metric", y="Score", hue="Model", data=df_melted, palette="Set2")
    
    plt.ylim([0.7, 1.02])
    plt.ylabel('Score', fontsize=12, fontweight='bold')
    plt.xlabel('Evaluation Metrics', fontsize=12, fontweight='bold')
    plt.title('Performance Comparison of Models', fontsize=14, fontweight='bold')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0., fontsize=11)
    
    # Annotate bars
    for p in ax.patches:
        height = p.get_height()
        if height > 0:
            ax.annotate(f'{height:.3f}',
                        (p.get_x() + p.get_width() / 2., height),
                        ha='center', va='bottom',
                        xytext=(0, 3),
                        textcoords='offset points',
                        fontsize=8)
                        
    plt.tight_layout()
    plt.savefig(save_path, dpi=300)
    plt.close()
    print(f"Comparison chart saved to {save_path}")

def main():
    csv_path = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\data\fake_job_postings.csv"
    results_dir = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\results"
    os.makedirs(results_dir, exist_ok=True)
    
    from data_loader import get_splits
    from baselines import train_ml_baselines
    from train import train_deep_models
    
    print("Step 1: Loading and splitting dataset...")
    train, val, test = get_splits(csv_path)
    
    import torch
    if not torch.cuda.is_available():
        print("CUDA is NOT available. Running on CPU. Downsampling dataset to speed up training...")
        from sklearn.model_selection import train_test_split
        train, _ = train_test_split(train, train_size=2600, stratify=train['fraudulent'], random_state=42)
        val, _ = train_test_split(val, train_size=380, stratify=val['fraudulent'], random_state=42)
        test, _ = train_test_split(test, train_size=750, stratify=test['fraudulent'], random_state=42)
        print(f"Downsampled size - Train: {len(train)}, Val: {len(val)}, Test: {len(test)}")
    
    all_results = {}
    
    # 1. TF-IDF + ML Baselines
    print("\n--- Training Machine Learning Baselines ---")
    ml_metrics = train_ml_baselines(train, test, results_dir=results_dir)
    
    # Reload ML models to compute predictions and probabilities for plotting
    with open(os.path.join(results_dir, "tfidf_vectorizer.pkl"), "rb") as f:
        vectorizer = joblib.load(f)
        
    X_test_tfidf = vectorizer.transform(test['combined_text'])
    y_test = test['fraudulent'].values
    
    for name in ["Logistic Regression", "Random Forest"]:
        model_filename = name.lower().replace(" ", "_") + ".pkl"
        with open(os.path.join(results_dir, model_filename), "rb") as f:
            model = joblib.load(f)
            
        preds = model.predict(X_test_tfidf)
        probs = model.predict_proba(X_test_tfidf)[:, 1]
        
        all_results[name] = {
            "accuracy": ml_metrics[name]["Accuracy"],
            "precision": ml_metrics[name]["Precision"],
            "recall": ml_metrics[name]["Recall"],
            "f1": ml_metrics[name]["F1-score"],
            "auc": ml_metrics[name]["ROC-AUC"],
            "preds": preds,
            "probs": probs,
            "y_true": y_test
        }
        
    # 2. PyTorch Deep Learning Models
    deep_models = {
        "LSTM": "lstm",
        "Standalone BERT": "bert_standalone",
        "Standalone RoBERTa": "roberta_standalone",
        "Proposed BERT-BiLSTM": "bert_bilstm"
    }
    
    for name, model_type in deep_models.items():
        print(f"\n--- Training {name} ---")
        # Run training (setting epochs=3 for transformer, 5 for LSTM)
        epochs = 5 if model_type == "lstm" else 3
        res = train_deep_models(
            train, val, test, 
            model_type=model_type, 
            epochs=epochs, 
            batch_size=32, 
            results_dir=results_dir
        )
        all_results[name] = res
        
    print("\n--- Generating Evaluation Charts ---")
    # Plot combined ROC
    plot_roc_curves(all_results, os.path.join(results_dir, "roc_curves.png"))
    
    # Plot proposed model confusion matrix
    plot_confusion_matrix(
        all_results["Proposed BERT-BiLSTM"]["y_true"], 
        all_results["Proposed BERT-BiLSTM"]["preds"], 
        os.path.join(results_dir, "confusion_matrix_proposed.png"),
        "Proposed BERT-BiLSTM Confusion Matrix"
    )
    
    # Plot LSTM confusion matrix for comparison
    plot_confusion_matrix(
        all_results["LSTM"]["y_true"], 
        all_results["LSTM"]["preds"], 
        os.path.join(results_dir, "confusion_matrix_lstm.png"),
        "Standard Bi-LSTM Confusion Matrix"
    )
    
    # Plot overall performance comparison
    plot_comparison_bar_chart(all_results, os.path.join(results_dir, "performance_comparison.png"))
    
    # Save a summary table to results/summary.csv
    summary_data = []
    for name, res in all_results.items():
        summary_data.append({
            "Model": name,
            "Accuracy": f"{res['accuracy']*100:.2f}%",
            "Precision": f"{res['precision']*100:.2f}%",
            "Recall": f"{res['recall']*100:.2f}%",
            "F1-score": f"{res['f1']*100:.2f}%",
            "ROC-AUC": f"{res['auc']*100:.2f}%"
        })
    df_summary = pd.DataFrame(summary_data)
    df_summary.to_csv(os.path.join(results_dir, "summary.csv"), index=False)
    
    print("\n=== All experiments finished! ===")
    print(df_summary)

if __name__ == "__main__":
    main()
