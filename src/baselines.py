import os
import joblib
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, roc_auc_score

def train_ml_baselines(train_df, test_df, results_dir="results"):
    os.makedirs(results_dir, exist_ok=True)
    
    print("Vectorizing text using TF-IDF...")
    # Clean text columns are in 'combined_text'
    X_train = train_df['combined_text']
    y_train = train_df['fraudulent']
    
    X_test = test_df['combined_text']
    y_test = test_df['fraudulent']
    
    # TF-IDF parameters
    vectorizer = TfidfVectorizer(max_features=5000, ngram_range=(1, 2), stop_words='english')
    X_train_tfidf = vectorizer.fit_transform(X_train)
    X_test_tfidf = vectorizer.transform(X_test)
    
    # Save vectorizer
    joblib.dump(vectorizer, os.path.join(results_dir, "tfidf_vectorizer.pkl"))
    
    models = {
        "Logistic Regression": LogisticRegression(class_weight="balanced", max_iter=1000, random_state=42),
        "Random Forest": RandomForestClassifier(class_weight="balanced", n_estimators=100, random_state=42)
    }
    
    metrics_summary = {}
    
    for name, model in models.items():
        print(f"Training {name}...")
        model.fit(X_train_tfidf, y_train)
        
        # Save model
        model_filename = name.lower().replace(" ", "_") + ".pkl"
        joblib.dump(model, os.path.join(results_dir, model_filename))
        
        # Predict
        preds = model.predict(X_test_tfidf)
        probs = model.predict_proba(X_test_tfidf)[:, 1]
        
        # Calculate metrics
        acc = accuracy_score(y_test, preds)
        prec = precision_score(y_test, preds, zero_division=0)
        rec = recall_score(y_test, preds)
        f1 = f1_score(y_test, preds)
        auc = roc_auc_score(y_test, probs)
        
        print(f"{name} Results - Acc: {acc:.4f}, Precision: {prec:.4f}, Recall: {rec:.4f}, F1: {f1:.4f}, AUC: {auc:.4f}")
        
        metrics_summary[name] = {
            "Accuracy": acc,
            "Precision": prec,
            "Recall": rec,
            "F1-score": f1,
            "ROC-AUC": auc
        }
        
    return metrics_summary

if __name__ == "__main__":
    from data_loader import get_splits
    csv_path = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\data\fake_job_postings.csv"
    train, val, test = get_splits(csv_path)
    train_ml_baselines(train, test)
