import os
import re
import pandas as pd
from sklearn.model_selection import train_test_split

def clean_text(text):
    if not isinstance(text, str):
        return ""
    # Remove HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    # Remove special characters and numbers (optional, keeping basic punctuation might help transformers)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def load_and_preprocess_data(csv_path):
    print("Loading dataset...")
    df = pd.read_csv(csv_path)
    
    # Fill NaN values in text columns with empty strings
    text_cols = ['title', 'company_profile', 'description', 'requirements', 'benefits']
    for col in text_cols:
        df[col] = df[col].fillna("")
        df[col] = df[col].apply(clean_text)
    
    # Combine text fields to form a single input string for NLP models
    print("Combining text features...")
    df['combined_text'] = (
        "Title: " + df['title'] + "\n" +
        "Company Profile: " + df['company_profile'] + "\n" +
        "Description: " + df['description'] + "\n" +
        "Requirements: " + df['requirements'] + "\n" +
        "Benefits: " + df['benefits']
    )
    
    # Filter out columns we don't need for the text model, but keep target
    df = df[['combined_text', 'fraudulent']].copy()
    
    # Clean up empty texts if any
    df = df[df['combined_text'].str.strip() != ""]
    
    print(f"Dataset size: {len(df)}")
    print(f"Class distribution:\n{df['fraudulent'].value_counts(normalize=True)}")
    
    return df

def get_splits(csv_path, test_size=0.2, val_size=0.1, random_state=42):
    df = load_and_preprocess_data(csv_path)
    
    # First split train + val and test
    train_val, test = train_test_split(
        df, 
        test_size=test_size, 
        stratify=df['fraudulent'], 
        random_state=random_state
    )
    
    # Then split train and val
    # val_size relative to train_val should equal the original val_size
    relative_val_size = val_size / (1.0 - test_size)
    train, val = train_test_split(
        train_val, 
        test_size=relative_val_size, 
        stratify=train_val['fraudulent'], 
        random_state=random_state
    )
    
    print(f"Train size: {len(train)}, Val size: {len(val)}, Test size: {len(test)}")
    return train, val, test

if __name__ == "__main__":
    csv_path = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\data\fake_job_postings.csv"
    train, val, test = get_splits(csv_path)
    print("Sample text from train:")
    print(train.iloc[0]['combined_text'][:500])
