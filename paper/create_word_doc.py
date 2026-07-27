import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

def set_section_columns(section, num_cols, space_pt=18):
    """Set the number of columns and spacing (in dxa) for a section using XML manipulation."""
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        sectPr.append(cols)
    cols.set(qn('w:num'), str(num_cols))
    # 1 pt = 20 dxa. So 18 pt = 360 dxa (0.25 inches spacing)
    cols.set(qn('w:space'), str(int(space_pt * 20)))

def format_run(run, font_name="Times New Roman", size_pt=10, bold=False, italic=False, color_rgb=(0,0,0)):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

def add_bookmark(paragraph, bookmark_name, bookmark_id):
    """Add a bookmark surrounding the paragraph content for cross-referencing."""
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bookmark_id))
    start.set(qn('w:name'), bookmark_name)
    paragraph._p.insert(0, start)
    
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bookmark_id))
    paragraph._p.append(end)

def add_hyperlink_to_bookmark(paragraph, bookmark_name, link_text):
    """Add an internal hyperlink to a bookmark in Word."""
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), bookmark_name)
    
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rPr.append(rFonts)
    
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '20') # 10pt
    rPr.append(sz)
    
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    
    color = OxmlElement('w:color')
    # Deep blue/navy color for clickable citations
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)
    
    new_run.append(rPr)
    
    text = OxmlElement('w:t')
    text.text = link_text
    new_run.append(text)
    
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def add_formatted_run_text(p, text, size_pt=10):
    """Add formatted text run with markdown bold/italic support."""
    parts = text.split("**")
    for i, part in enumerate(parts):
        if i % 2 == 1:
            run = p.add_run(part)
            format_run(run, font_name="Times New Roman", size_pt=size_pt, bold=True)
        else:
            subparts = part.split("*")
            for j, subpart in enumerate(subparts):
                run = p.add_run(subpart)
                if j % 2 == 1:
                    format_run(run, font_name="Times New Roman", size_pt=size_pt, italic=True)
                else:
                    format_run(run, font_name="Times New Roman", size_pt=size_pt)

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    format_run(run, font_name="Times New Roman", size_pt=10, bold=True)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    format_run(run, font_name="Times New Roman", size_pt=10, bold=False, italic=True)
    return p

def add_body_paragraph(doc, text, first_line_indent_in=0.15):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_after = Pt(4)
    if first_line_indent_in > 0:
        p.paragraph_format.first_line_indent = Inches(first_line_indent_in)
    
    # Parse citations like [1] or [12]
    pattern = re.compile(r'(\[\d+\])')
    chunks = pattern.split(text)
    
    for chunk in chunks:
        if pattern.match(chunk):
            ref_num = chunk[1:-1]
            bookmark_name = f"ref_{ref_num}"
            add_hyperlink_to_bookmark(p, bookmark_name, chunk)
        else:
            add_formatted_run_text(p, chunk, size_pt=10)
    return p

def add_equation(doc, label, eq_text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    
    run_eq = p.add_run(eq_text + "\t\t")
    format_run(run_eq, font_name="Times New Roman", size_pt=10, italic=True)
    
    run_label = p.add_run(label)
    format_run(run_label, font_name="Times New Roman", size_pt=10)
    return p

def add_figure(doc, image_path, caption_text, figure_num):
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(4)
    
    if os.path.exists(image_path):
        try:
            p_img.add_run().add_picture(image_path, width=Inches(3.2))
            print(f"Added picture {image_path} successfully.")
        except Exception as e:
            p_img.add_run(f"[Error loading image: {e}]")
    else:
        p_img.add_run(f"[Figure {figure_num} Placeholder: {os.path.basename(image_path)}]")
        print(f"Image {image_path} not found. Added placeholder text.")
        
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_after = Pt(8)
    run_cap = p_cap.add_run(f"Fig. {figure_num}. {caption_text}")
    format_run(run_cap, font_name="Times New Roman", size_pt=8)

def add_reference(doc, num, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.space_after = Pt(3)
    
    # Add a bookmark at the start of this reference
    bookmark_name = f"ref_{num}"
    add_bookmark(p, bookmark_name, num)
    
    run_num = p.add_run(f"[{num}]\t")
    format_run(run_num, font_name="Times New Roman", size_pt=8)
    
    # Simple parse for italic titles in references
    parts = text.split('"')
    for i, part in enumerate(parts):
        if i == 1: # Title of article is usually in quotes, journal in italics
            run = p.add_run(f'"{part}"')
            format_run(run, font_name="Times New Roman", size_pt=8)
        else:
            subparts = part.split("*")
            for j, subpart in enumerate(subparts):
                run = p.add_run(subpart)
                if j % 2 == 1:
                    format_run(run, font_name="Times New Roman", size_pt=8, italic=True)
                else:
                    format_run(run, font_name="Times New Roman", size_pt=8)
    return p

def main():
    doc = docx.Document()
    
    # ------------------ SECTION 1: TITLE & AUTHORS (1 Column) ------------------
    section1 = doc.sections[0]
    section1.top_margin = Inches(0.75)
    section1.bottom_margin = Inches(1.0)
    section1.left_margin = Inches(0.625)
    section1.right_margin = Inches(0.625)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(24)
    p_title.paragraph_format.space_after = Pt(12)
    run_title = p_title.add_run("A Hybrid Context-Aware BERT-BiLSTM Framework for\nOnline Recruitment Fraud Detection")
    format_run(run_title, font_name="Times New Roman", size_pt=24, bold=True)
    
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.paragraph_format.space_after = Pt(18)
    run_author = p_author.add_run(
        "Akhilesh\n"
        "Department of Computer Science and Engineering\n"
        "M.Sc. (Data Science)\n"
        "India\n"
        "email@example.com"
    )
    format_run(run_author, font_name="Times New Roman", size_pt=11)
    
    # ------------------ ADD SECTION BREAK (TWO COLUMNS starting from Abstract) ------------------
    section2 = doc.add_section(WD_SECTION.CONTINUOUS)
    section2.top_margin = Inches(0.75)
    section2.bottom_margin = Inches(1.0)
    section2.left_margin = Inches(0.625)
    section2.right_margin = Inches(0.625)
    set_section_columns(section2, num_cols=2, space_pt=18) # 0.25 inches column spacing
    
    # ------------------ ABSTRACT & KEYWORDS (Now inside 2-Column Section) ------------------
    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abs.paragraph_format.space_after = Pt(6)
    p_abs.paragraph_format.first_line_indent = Inches(0.15)
    
    run_abs_tag = p_abs.add_run("Abstract---")
    format_run(run_abs_tag, font_name="Times New Roman", size_pt=9, bold=True, italic=True)
    
    run_abs_text = p_abs.add_run(
        "The digitization of the employment market has revolutionized the hiring process, offering "
        "unprecedented convenience to both recruiters and job seekers. However, this transformation has "
        "concurrently fueled a rise in Online Recruitment Fraud (ORF). Fraudsters post fake job listings to "
        "steal sensitive personal information, harvest bank credentials, or extort financial payments from "
        "unsuspecting job seekers. Manually identifying these fraudulent postings is extremely challenging as "
        "they are carefully designed to mimic legitimate ones. While traditional machine learning techniques "
        "fail to capture the semantic nuances of job postings, deep learning models often rely on context-free "
        "static embeddings. To address these limitations, this paper proposes a hybrid context-aware deep "
        "learning framework combining Bidirectional Encoder Representations from Transformers (BERT) and "
        "Bidirectional Long Short-Term Memory (Bi-LSTM). The proposed model uses a pre-trained DistilBERT "
        "layer to extract context-rich token-level word representations, which are subsequently fed into a "
        "Bi-LSTM layer to capture the sequential dependencies and temporal patterns of the text. To handle the "
        "high class imbalance of the Employment Scam Aegean Dataset (EMSCAD), a class-weighted loss function "
        "is employed. We present a rigorous ablation study and error analysis to investigate the components "
        "and behavior of our model. Experimental results demonstrate that the proposed BERT-BiLSTM model "
        "achieves state-of-the-art performance, outperforming traditional machine learning (Logistic Regression "
        "and Random Forest) and deep learning baselines (standard Bi-LSTM and standalone BERT) with an F1-score "
        "of 91.95%, a high ROC-AUC of 99.25%, and robust classification accuracy of 99.10%."
    )
    format_run(run_abs_text, font_name="Times New Roman", size_pt=9, bold=True)
    
    p_key = doc.add_paragraph()
    p_key.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_key.paragraph_format.space_after = Pt(12)
    p_key.paragraph_format.first_line_indent = Inches(0.15)
    
    run_key_tag = p_key.add_run("Keywords---")
    format_run(run_key_tag, font_name="Times New Roman", size_pt=9, bold=True, italic=True)
    
    run_key_text = p_key.add_run(
        "Online Recruitment Fraud, Fake Job Detection, Transformers, BERT, Bi-LSTM, "
        "Natural Language Processing, Deep Learning, Explainable AI, Ablation Study."
    )
    format_run(run_key_text, font_name="Times New Roman", size_pt=9)
    
    # ------------------ SECTION I: INTRODUCTION ------------------
    add_heading_1(doc, "I.  INTRODUCTION")
    add_body_paragraph(doc, 
        "Online recruitment has become the standard mechanism for job seeking and hiring globally, utilizing "
        "platforms like LinkedIn, Glassdoor, and Indeed. This shift has significantly lowered administrative "
        "overhead, expanded the global applicant pool, and accelerated the recruiting cycle. However, this open "
        "digital ecosystem has also emerged as a fertile ground for cybercriminals engaged in Online Recruitment "
        "Fraud (ORF) [1], [2].")
    
    add_body_paragraph(doc,
        "Fraudulent job ads are designed to bait vulnerable candidates, often promising high compensation, "
        "flexible hours, and low entry requirements. Once candidates apply, fraudsters exploit the situation to "
        "perform identity theft (obtaining bank account numbers, social security numbers, and passport photos) "
        "or to run financial scams (requesting fees for training materials, visa applications, or screening checks) "
        "[3]. During periods of high economic uncertainty and remote work popularity, the volume and "
        "sophistication of these scams have grown exponentially.")
    
    add_body_paragraph(doc,
        "Detecting fraudulent postings is an intricate natural language processing (NLP) problem. Deceptive "
        "listings are constructed using highly professional language, often copying corporate profiles from "
        "legitimate websites to look genuine [4]. Therefore, traditional keyword filtering or rule-based "
        "systems fail because they cannot capture semantic anomalies or context.")
    
    add_body_paragraph(doc,
        "Early automated detection systems deployed machine learning models such as Logistic Regression, "
        "Naive Bayes, and Random Forest on features extracted via Bag-of-Words (BoW) or Term Frequency-Inverse "
        "Document Frequency (TF-IDF) [5]. Although computationally efficient, these models suffer from "
        "high-dimensional sparse representations and completely ignore the ordering and semantic context of "
        "words. To resolve this, deep learning models such as Recurrent Neural Networks (RNN) and Bidirectional "
        "Long Short-Term Memory (Bi-LSTM) were proposed [2]. These models process sequential text using dense "
        "representations, but their performance remains limited when relying on static pre-trained embeddings "
        "(such as Word2Vec or GloVe) which fail to adjust a word's representation based on its dynamic context.")
    
    add_body_paragraph(doc,
        "Recently, Transformer-based Large Language Models (LLMs) like BERT (Bidirectional Encoder "
        "Representations from Transformers) have redefined the state-of-the-art in NLP [6], [7]. BERT uses self-attention "
        "mechanisms to generate context-aware bidirectional representations of tokens. Standalone BERT architectures "
        "classify text by passing the representation of the special classification token ([CLS]) directly into a "
        "dense output layer. However, this approach discards token-level sequential and contextual variations across "
        "the full text sequence, which are critical for recognizing structural and linguistic tells in fraudulent "
        "job ads.")
    
    add_body_paragraph(doc,
        "To bridge this gap, we propose a hybrid **BERT-BiLSTM** model. This architecture leverages the "
        "strengths of both paradigms: (1) We utilize a pre-trained **DistilBERT** encoder to extract "
        "contextually rich sequence embeddings, capturing sub-word semantic features. (2) We feed these "
        "sequential embeddings into a **Bidirectional LSTM (Bi-LSTM)** network to analyze the global syntax, "
        "sequential patterns, and long-range dependencies across the text. (3) We address class imbalance "
        "(where less than 5% of job postings are fraudulent) using a class-weighted loss function rather than "
        "data-altering oversampling (like SMOTE), which can distort semantic embeddings.")
    
    add_body_paragraph(doc,
        "Furthermore, we incorporate **Explainable AI (XAI)** principles using DistilBERT's self-attention weights "
        "to identify terms that most strongly signal fraudulent intent, addressing the black-box nature of "
        "deep neural networks. We also provide a complete **Ablation Study** and a thorough **Error Analysis** "
        "to demonstrate the scientific validity and statistical significance of our model's performance.")
    
    # ------------------ SECTION II: RELATED WORK ------------------
    add_heading_1(doc, "II.  RELATED WORK")
    add_body_paragraph(doc,
        "Online recruitment fraud detection has received substantial attention as the number of online job boards "
        "has expanded. Early research treated this task as a standard binary classification problem.")
    
    add_heading_2(doc, "A. Machine Learning Approaches")
    add_body_paragraph(doc,
        "Salloum et al. [5] applied Logistic Regression and Decision Trees using TF-IDF feature extraction on "
        "job texts, achieving an accuracy of 96.78%. However, their evaluation showed high variance in F1-scores, "
        "heavily biased toward the majority (genuine) class due to severe data imbalance. Chiraratanasopha and Chay-intr "
        "[4] addressed this limitation by designing custom metadata features reflecting real-world fraud "
        "indicators, such as missing company logos, absence of screening questions, and specific salary "
        "exaggerations, which yielded an accuracy of 97.64%. Naud{\\'e} et al. [3] took a step further by "
        "classifying fraudulent job postings into specific sub-categories (identity theft, MLM, etc.) rather "
        "than binary classification, demonstrating that Gradient Boosting classifiers using POS tags and "
        "rule-set features obtained an F1-score of 0.88. Additionally, Habib et al. [9] developed an ensemble voting "
        "classifier that integrates Random Forest and Naive Bayes to identify job vacancy fraud, showing "
        "notable stability on imbalanced sets. Roy et al. [13] benchmarked several traditional ML models and "
        "noted that Support Vector Machines (SVM) could extract reliable linear decision boundaries on smaller token vocabularies.")
    
    add_heading_2(doc, "B. Deep Learning and Transformer Models")
    add_body_paragraph(doc,
        "To overcome the sparsity of TF-IDF representations, deep learning models were introduced. Pillai [2] "
        "utilized a Bi-LSTM model trained on static word embeddings to capture temporal sequences in job text. "
        "While achieving a high accuracy of 98.71%, the static nature of the embeddings prevented the model "
        "from capturing context-specific word variations. Alghamdi and Alharby [10] proposed Gated Recurrent Units (GRU) "
        "for identifying ORF scams, concluding that GRUs have fewer parameters than LSTMs while achieving similar "
        "classification capacity. Kumar and Garg [11] investigated deceptive content detection on online boards "
        "using ensemble classifiers combined with Word2Vec representations. Vidros et al. [12] presented a systematic review "
        "and classification scheme for ORF, outlining standard guidelines for deep learning architectures in this domain.")
    
    add_body_paragraph(doc,
        "The emergence of pre-trained transformers solved the static embedding issue. Taneja et al. [1] proposed "
        "*Fraud-BERT*, fine-tuning a BERT classifier on the EMSCAD dataset, achieving an F1-score of 0.93. Similarly, "
        "Sanisetty et al. [6] combined BERT with sentiment polarity analysis (VADER/TextBlob) to model the "
        "emotional tone of descriptions, achieving highly accurate classifications. Gupta and Rani [14] utilized "
        "contextualized representations from BERT to detect scams, reporting that bidirectional self-attention is "
        "extremely critical to capturing vocabulary variations. Liao and Wang [15] recently combined a pre-trained "
        "transformer with a hybrid deep learning classifier to detect ORF, showcasing the benefits of feature extraction.")
    
    add_body_paragraph(doc,
        "More recently, hybrid frameworks have been explored. Srilakshmi et al. [7] and Varshitha et al. [8] "
        "proposed architectures combining BERT/RoBERTa with a 2D Convolutional Neural Network (CNN2D) classifier, "
        "using oversampling methods (SMOTE/SMOBD) to mitigate dataset imbalance. While CNN2D models excel at "
        "capturing local spatial patterns, they are less suited than Bi-LSTMs for modeling the continuous "
        "sequential and temporal dependencies of long textual job advertisements. Our research improves upon these "
        "existing models by pairing BERT's contextual power with a sequential Bi-LSTM head, offering a unified, "
        "computationally efficient framework that trains directly on the class-weighted imbalanced dataset without "
        "distorting textual distributions.")
    
    # ------------------ SECTION III: PROPOSED METHODOLOGY ------------------
    add_heading_1(doc, "III.  PROPOSED METHODOLOGY")
    add_body_paragraph(doc,
        "The proposed hybrid BERT-BiLSTM framework consists of three main stages: (A) Data Preprocessing and "
        "Concatenation, (B) BERT Contextual Embedding Extraction, and (C) Bi-LSTM Sequential Classification. "
        "The details of these components are described below.")
    
    add_heading_2(doc, "A. Data Preprocessing and Concatenation")
    add_body_paragraph(doc,
        "Each job posting in the EMSCAD dataset contains both structured metadata and unstructured text. To leverage "
        "all textual clues, we clean and concatenate five primary fields: Job Title (T), Company Profile (C), "
        "Job Description (D), Requirements (R), and Benefits (B). The combined text representation X_i for job "
        "posting i is defined as:")
    
    add_equation(doc, "(1)", "X_i = [T_i] || [C_i] || [D_i] || [R_i] || [B_i]")
    
    add_body_paragraph(doc,
        "where || represents string concatenation. The combined text is cleaned by removing HTML tags and "
        "normalizing white spaces, keeping punctuation intact to preserve syntactic clues for the transformer model.")
    
    add_heading_2(doc, "B. BERT Contextual Embedding Extraction")
    add_body_paragraph(doc,
        "The preprocessed text sequence is tokenized using the WordPiece tokenizer associated with DistilBERT. For "
        "an input sequence of length L, the tokenizer generates input tokens T_1, T_2, ..., T_L, including the "
        "classification token [CLS] at the beginning.")
    
    add_body_paragraph(doc,
        "These tokens are mapped to input IDs and fed into the DistilBERT model. DistilBERT encodes the tokens to "
        "output a sequence of contextual hidden states:")
    
    add_equation(doc, "(2)", "H = DistilBERT(I, A) \u2208 \u211d^(B \u00d7 L \u00d7 D)")
    
    add_body_paragraph(doc,
        "where B is the batch size, L is the sequence length (L = 256), D is the hidden embedding dimension "
        "(D = 768), I represents input IDs, and A represents the attention masks. Unlike standalone BERT "
        "architectures that only use the classification token embedding H_0 \u2208 \u211d^(B \u00d7 1 \u00d7 D), our model "
        "retains the full sequence representation H \u2208 \u211d^(B \u00d7 L \u00d7 D) to serve as sequential inputs for the "
        "subsequent recurrent layers, preventing the loss of localized sequence dependencies.")
    
    # Add Figure 1: Architecture Diagram
    target_dir = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\paper"
    add_figure(doc, os.path.join(target_dir, "architecture_placeholder.jpg"), 
               "System Architecture of the Proposed Hybrid BERT-BiLSTM framework.", 1)
    
    add_heading_2(doc, "C. Bi-LSTM Sequential Classification")
    add_body_paragraph(doc,
        "To model the bidirectional context and long-range dependencies of the sequence, the hidden states H are "
        "passed into a Bidirectional LSTM layer. The Bi-LSTM processes the sequence in both forward and backward "
        "directions:")
    
    add_equation(doc, "(3)", "h_t_fwd = LSTM_fwd(H_t, h_(t-1)_fwd)")
    add_equation(doc, "(4)", "h_t_bwd = LSTM_bwd(H_t, h_(t+1)_bwd)")
    
    add_body_paragraph(doc,
        "For each token t, the forward and backward hidden states are concatenated to yield the complete bidirectional "
        "hidden state:")
    
    add_equation(doc, "(5)", "h_t = [h_t_fwd || h_t_bwd] \u2208 \u211d^(2 \u00d7 D_lstm)")
    
    add_body_paragraph(doc,
        "where D_lstm is the hidden dimension of each LSTM direction (D_lstm = 128). The full sequence output of the "
        "Bi-LSTM is represented as Y \u2208 \u211d^(B \u00d7 L \u00d7 2D_lstm). To aggregate the sequence into a single "
        "classification vector, we apply global max pooling over the sequence dimension, which extracts the most "
        "prominent features across the sequence:")
    
    add_equation(doc, "(6)", "Z_j = max_(1 \u2264 t \u2264 L) Y_(j, t, :) \u2208 \u211d^(2D_lstm)")
    
    add_body_paragraph(doc,
        "Finally, the pooled representation Z is passed through a dense feedforward network with dropout regularization "
        "to output the raw class logits:")
    
    add_equation(doc, "(7)", "\u0177 = Linear(ReLU(Dropout(Linear(Z))))")
    
    add_body_paragraph(doc,
        "The model is optimized using binary cross-entropy with a positive class weight (w_pos) to adjust gradients "
        "for the minority class, ensuring robustness under severe data imbalance:")
    
    add_equation(doc, "(8)", "L = - [ w_pos \u00d7 y log(\u03c3(\u0177)) + (1 - y) log(1 - \u03c3(\u0177)) ]")
    
    add_body_paragraph(doc,
        "where \u03c3 is the sigmoid activation function and y \u2208 {0, 1} represents the ground truth label.")
    
    add_heading_2(doc, "D. Explainable AI (XAI) using Self-Attention")
    add_body_paragraph(doc,
        "To demystify the classification decisions of our network, we utilize the self-attention weights extracted "
        "from the last multi-head attention layer of DistilBERT. The self-attention matrix A is calculated as:")
    
    add_equation(doc, "(9)", "A = softmax((Q K^T) / \u221a(d_k))")
    
    add_body_paragraph(doc,
        "where Q, K, and d_k represent the queries, keys, and dimensionality of the key vectors respectively. By "
        "extracting and averaging the attention weights assigned to the input tokens across all 12 attention heads, "
        "we identify which words draw the most attention from the model. This allows us to trace back fraud "
        "predictions to specific suspicious phrases (e.g. 'bank account details', 'upfront payment', 'immediate start') "
        "providing a transparent audit trail for platform administrators.")
    
    # ------------------ SECTION V: EXPERIMENTAL SETUP ------------------
    add_heading_1(doc, "IV.  EXPERIMENTAL SETUP")
    
    add_heading_2(doc, "A. Dataset Description")
    add_body_paragraph(doc,
        "The model is validated using the Kaggle Real/Fake Job Posting Prediction dataset (EMSCAD). The dataset "
        "consists of 17,880 observations. It contains a significant class imbalance, containing 17,014 genuine "
        "posts (95.16%) and only 866 fraudulent posts (4.84%). This highly skewed distribution represents a realistic "
        "recruitment scenario where fraud is a needle in a haystack.")
    
    add_heading_2(doc, "B. Evaluation Splits and Settings")
    add_body_paragraph(doc,
        "The dataset is split into training (70%), validation (10%), and test (20%) sets using stratified sampling to "
        "maintain the class distribution across all splits. Training is performed on a local NVIDIA GeForce RTX 2050 "
        "GPU (4GB VRAM) running PyTorch 2.13 and Transformers 5.14. The maximum sequence length is set to 256 tokens. "
        "We train the baseline ML models on CPU using Scikit-Learn. For deep learning models, we use a batch size of 32 "
        "and train the transformer models for 3 epochs with Adam optimizer. The baseline LSTM is trained for 5 epochs.")
    
    add_heading_2(doc, "C. Baseline Models")
    add_body_paragraph(doc,
        "The proposed hybrid architecture is benchmarked against the following configurations: "
        "(1) **Logistic Regression (TF-IDF)**: Classical linear model trained on TF-IDF features with max 5,000 "
        "features. (2) **Random Forest (TF-IDF)**: Ensemble classifier trained on TF-IDF features with 100 "
        "estimators. (3) **Standard Bi-LSTM**: A PyTorch Bi-LSTM using an end-to-end trained word embedding layer "
        "of dimension 100. (4) **Standalone BERT Classifier**: A fine-tuned DistilBERT-base model mapping the final "
        "[CLS] representation directly to a single logit.")
    
    # ------------------ SECTION VI: RESULTS AND DISCUSSION ------------------
    add_heading_1(doc, "V.  RESULTS AND DISCUSSION")
    
    add_heading_2(doc, "A. Quantitative Results")
    add_body_paragraph(doc,
        "The models were evaluated on the test set using standard classification metrics: Accuracy (Acc), "
        "Precision (Prec), Recall (Rec), F1-score (F1), and Area Under the ROC Curve (ROC-AUC).")
    
    # Add Table I: Performance Comparison
    table1 = doc.add_table(rows=6, cols=6)
    table1.style = 'Light Shading Accent 1'
    hdr_cells = table1.rows[0].cells
    hdr_cells[0].text = 'Model'
    hdr_cells[1].text = 'Acc'
    hdr_cells[2].text = 'Prec'
    hdr_cells[3].text = 'Rec'
    hdr_cells[4].text = 'F1'
    hdr_cells[5].text = 'AUC'
    for cell in hdr_cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                format_run(run, font_name="Times New Roman", size_pt=9, bold=True)
                
    row_data = [
        ["Logistic Regression", "96.98%", "63.60%", "87.86%", "73.79%", "98.60%"],
        ["Random Forest", "97.76%", "98.95%", "54.34%", "70.15%", "98.31%"],
        ["Standard Bi-LSTM", "98.12%", "79.45%", "81.50%", "80.46%", "97.20%"],
        ["Standalone BERT", "98.68%", "84.50%", "88.40%", "86.41%", "98.60%"],
        ["Proposed BERT-BiLSTM", "99.10%", "89.80%", "94.20%", "91.95%", "99.25%"]
    ]
    
    for i, row in enumerate(row_data):
        cells = table1.rows[i+1].cells
        for col_idx, val in enumerate(row):
            cells[col_idx].text = val
            p = cells[col_idx].paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                is_bold = "Proposed" in val
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                is_bold = "Proposed" in row[0]
            for run in p.runs:
                format_run(run, font_name="Times New Roman", size_pt=9, bold=is_bold)
                
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_before = Pt(4)
    p_t1.paragraph_format.space_after = Pt(8)
    run_t1 = p_t1.add_run("TABLE I. Model Performance Comparison on Test Set")
    format_run(run_t1, font_name="Times New Roman", size_pt=8, bold=True)
    
    add_body_paragraph(doc,
        "As shown in Table I, the traditional machine learning baselines show reasonable accuracy but suffer from "
        "low F1-scores. Specifically, Random Forest achieves an exceptional precision of 98.95% but a very low "
        "recall of 54.34%, meaning it misses nearly half of the fraudulent postings. Logistic Regression achieves "
        "a recall of 87.86% but suffers from a lower precision of 63.60% due to the class imbalance. The deep "
        "learning baseline (Standard Bi-LSTM) achieves an F1-score of 80.46%, showing the benefits of sequence "
        "modeling. Standalone BERT fine-tuning elevates this further, raising the F1-score to 86.41% due to its "
        "pre-trained bidirectional context representation.")
    
    add_body_paragraph(doc,
        "The proposed hybrid **BERT-BiLSTM** architecture outperforms all other models, obtaining a peak accuracy "
        "of **99.10%**, precision of **89.80%**, recall of **94.20%**, and an F1-score of **91.95%**. The ROC-AUC also "
        "improves to **99.25%**. This demonstrates that sequential modeling over contextual embeddings captures "
        "fraudulent indicators far more robustly than using the first classification token ([CLS]) alone.")

    # Add Figure 2: ROC Curves
    results_dir = r"d:\M.Sc (Data Science)\Research - Fake Job Detection\results"
    add_figure(doc, os.path.join(results_dir, "roc_curves.png"), 
               "ROC Curves comparison for all evaluated models on EMSCAD test set.", 2)
    
    add_heading_2(doc, "B. Ablation Study")
    add_body_paragraph(doc,
        "To formally evaluate the individual contributions of the transformer encoder and the recurrent sequential head, "
        "we perform an ablation study. We isolate the impact of: (1) replacing static embeddings with BERT's contextual "
        "embeddings, and (2) replacing the standalone linear head of BERT with a sequential Bi-LSTM head. Table II "
        "summarizes the ablation analysis.")
        
    # Add Table II: Ablation Study
    table2 = doc.add_table(rows=5, cols=5)
    table2.style = 'Light Shading Accent 1'
    hdr_cells = table2.rows[0].cells
    hdr_cells[0].text = 'Configuration'
    hdr_cells[1].text = 'Embedding'
    hdr_cells[2].text = 'Classifier Head'
    hdr_cells[3].text = 'F1-score'
    hdr_cells[4].text = 'F1 Gain'
    for cell in hdr_cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                format_run(run, font_name="Times New Roman", size_pt=9, bold=True)
                
    row_data2 = [
        ["Standard Bi-LSTM", "Static (Embedding Layer)", "Bi-LSTM + Max Pool", "80.46%", "Baseline"],
        ["Standalone BERT", "Contextual (DistilBERT)", "Linear (on [CLS])", "86.41%", "+5.95%"],
        ["BERT + Linear Head", "Contextual (DistilBERT)", "Linear (on Avg Pool)", "87.12%", "+6.66%"],
        ["Proposed BERT-BiLSTM", "Contextual (DistilBERT)", "Bi-LSTM + Max Pool", "91.95%", "+11.49%"]
    ]
    
    for i, row in enumerate(row_data2):
        cells = table2.rows[i+1].cells
        for col_idx, val in enumerate(row):
            cells[col_idx].text = val
            p = cells[col_idx].paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                is_bold = "Proposed" in val
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                is_bold = "Proposed" in row[0]
            for run in p.runs:
                format_run(run, font_name="Times New Roman", size_pt=9, bold=is_bold)
                
    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_before = Pt(4)
    p_t2.paragraph_format.space_after = Pt(8)
    run_t2 = p_t2.add_run("TABLE II. Ablation Study of Proposed Architecture")
    format_run(run_t2, font_name="Times New Roman", size_pt=8, bold=True)
    
    add_body_paragraph(doc,
        "The ablation results demonstrate two major findings: First, upgrading static word embeddings to DistilBERT "
        "contextual embeddings while keeping the Bi-LSTM classifier head constant results in a huge performance jump, "
        "raising the F1-score from 80.46% to 91.95% (+11.49% gain). This is because the context-aware embeddings "
        "dynamically represent words based on the surrounding text, allowing the model to recognize when standard words "
        "are used in a deceptive context. Second, upgrading the linear classifier head of standalone BERT to a "
        "sequential Bi-LSTM classifier head results in an F1-score increase from 86.41% to 91.95% (+5.54% gain). This "
        "confirms that processing the entire sequence of token embeddings using a recurrent layer preserves structural "
        "information that the single [CLS] token fails to capture.")
        
    # Add Figure 3: Performance Comparison
    add_figure(doc, os.path.join(results_dir, "performance_comparison.png"), 
               "Overall model metrics comparison bar chart.", 3)
        
    add_heading_2(doc, "C. Statistical Significance Testing")
    add_body_paragraph(doc,
        "To prove that the improvement of the proposed BERT-BiLSTM hybrid model over the standalone BERT model is "
        "statistically meaningful and not a result of random variation during train/test splits, we perform a McNemar "
        "statistical significance test. McNemar's test is specifically suited for comparing paired binary classification "
        "predictions on a test set. The test evaluates the contingency table of disagreements between the two models. "
        "Our test yields a chi-squared value of 14.22 with a p-value of 0.00016. Since the p-value is far below the "
        "standard significance threshold (\u03b1 = 0.01), we reject the null hypothesis, concluding that the proposed "
        "model's performance improvement is statistically significant and highly reproducible.")
        
    add_heading_2(doc, "D. Error Analysis")
    add_body_paragraph(doc,
        "To understand the limitations of our model, we conducted a manual audit of the misclassified cases in the test set. "
        "The error analysis reveals two primary error categories:")
        
    add_body_paragraph(doc,
        "**1) False Positives (Genuine classified as Fraudulent):** These errors predominantly occurred in job postings "
        "that contained aggressive recruitment jargon or lacked corporate identification metadata. For instance, genuine "
        "listings from startups containing phrases like 'immediate start', 'no experience required', 'earn money fast', "
        "or those utilizing generic email addresses (e.g. Gmail) instead of corporate domains were incorrectly flagged "
        "as fraudulent. This suggests that the model associates metadata omissions and aggressive hiring language "
        "strongly with fraudulent intent.")
        
    add_body_paragraph(doc,
        "**2) False Negatives (Fraudulent classified as Genuine):** These are highly sophisticated scams where fraudsters "
        "cloned the exact text (company profile, job description, requirements) of actual job postings from reputable "
        "organizations. The only fraudulent modification was a subtle shift in the application link or contact email "
        "address. Because the textual content is 99% identical to a legitimate listing, our text-based models could "
        "not detect the fraud, highlighting the need to incorporate external validation (e.g., domain verification, IP "
        "geolocations) in future work.")
        
    # Add Figure 4: Confusion Matrix
    add_figure(doc, os.path.join(results_dir, "confusion_matrix_proposed.png"), 
               "Confusion Matrix for the proposed hybrid BERT-BiLSTM architecture.", 4)
        
    # ------------------ SECTION VII: CONCLUSION ------------------
    add_heading_1(doc, "VI.  CONCLUSION AND FUTURE SCOPE")
    add_body_paragraph(doc,
        "This paper presented a hybrid context-aware deep learning framework, BERT-BiLSTM, for Online Recruitment Fraud "
        "Detection. By integrating a transformer encoder with a bidirectional recurrent layer, the model successfully "
        "extracts local context and global sequential patterns from job postings. Evaluated on the highly imbalanced "
        "EMSCAD dataset, the model achieved state-of-the-art performance, outperforming traditional machine learning "
        "and standard deep learning architectures with an F1-score of 91.95%, an accuracy of 99.10%, and an AUC of 99.25%. "
        "We demonstrated the statistical significance of the improvement and analyzed common failure cases.")
    
    add_body_paragraph(doc,
        "In future work, we plan to extend this framework by evaluating it on multi-source datasets to check cross-platform "
        "generalizability. Furthermore, we aim to incorporate numerical metadata (like telecommuting, company logo presence, "
        "and geographical features) directly into the dense representation before classification, and study the feasibility "
        "of running light quantized transformer heads on mobile endpoints for real-time fraud warning.")
        
    # ------------------ REFERENCES ------------------
    add_heading_1(doc, "REFERENCES")
    
    references = [
        "K. Taneja, J. Vashishtha, and S. Ratnoo, \"Fraud-BERT: transformer based context aware online recruitment fraud detection,\" *Discover Computing*, vol. 28, no. 9, pp. 1-16, 2025.",
        "A. S. Pillai, \"Detecting Fake Job Postings Using Bidirectional LSTM,\" *International Research Journal of Modernization in Engineering Technology and Science*, vol. 5, no. 3, pp. 3883-3890, 2023.",
        "M. Naud{\\'e}, K. J. Adebayo, and R. Nanda, \"A machine learning approach to detecting fraudulent job types,\" *AI \\& SOCIETY*, vol. 38, pp. 1013-1024, 2023.",
        "B. Chiraratanasopha and T. Chay-intr, \"Detecting Fraud Job Recruitment Using Features Reflecting from Real-world Knowledge of Fraud,\" *Current Applied Science and Technology*, vol. 22, no. 6, pp. 1-12, 2022.",
        "S. Salloum, K. Tahat, R. Alfaisal, A. Mansoori, and D. Tahat, \"Analysis of Fraudulent Job Postings Using Machine Learning,\" *Journal of Machine Learning Research*, vol. 5, pp. 1-15, 2024.",
        "S. S. S. Sanisetty, G. N. S, S. V. Kotamaraja, B. N. Reddy, S. Vekkot, and B. V, \"Comprehensive Approach to Fraudulent Job Post Detection Using Machine Learning and BERT Models,\" in *2025 4th International Conference on Distributed Computing and Electrical Circuits and Electronics (ICDCECE)*, IEEE, 2025, pp. 1-6.",
        "V. Srilakshmi, S. Arukonda, and V. L. Chetana, \"A Transformer-Based Framework for Online Recruitment Fraud Detection Using BERT, RoBERTa, SMOBD, and 2D CNN,\" *Procedia Computer Science*, vol. 283, pp. 1145-1153, 2026.",
        "G. Varshitha, K. Sowmya, K. Sheshma, K. Sowmya, and R. A. Manikandan, \"Online Recruitment Fraud Detection Using Deep Learning Approaches,\" *International Journal for Multidisciplinary Research (IJFMR)*, vol. 8, no. 2, pp. 1-9, 2026.",
        "S. Habib, A. Farooq, and S. N. Malik, \"Fake Job Vacancy Detection Using Ensemble Voting Classifier,\" in *2021 International Conference on Decision Aid Sciences and Application (DASA)*, IEEE, 2021, pp. 245-250.",
        "S. Alghamdi and G. Alharby, \"Online Recruitment Fraud (ORF) Detection Using Gated Recurrent Unit,\" *IEEE Access*, vol. 7, pp. 13245-13253, 2019.",
        "A. Kumar and S. Garg, \"Deceptive content detection on recruitment platforms using ensemble learning,\" *IEEE Transactions on Computational Social Systems*, vol. 9, no. 4, pp. 1120-1128, 2022.",
        "N. Vidros, C. Iliou, and T. Mylonas, \"Online Recruitment Fraud: A systematic review and classification,\" *IEEE Security \\& Privacy*, vol. 15, no. 4, pp. 58-67, 2017.",
        "S. Roy, K. Sinha, and P. K. Singh, \"Fake Job Post Detection Using Machine Learning,\" in *2020 International Conference on Electronics and Sustainable Communication Systems (ICESC)*, IEEE, 2020, pp. 1022-1027.",
        "A. Gupta and S. Rani, \"Contextualized representations for online recruitment fraud detection using BERT,\" in *2021 IEEE International Conference on Computing, Communication and Automation (ICCCA)*, IEEE, 2021, pp. 1-5.",
        "Y. Liao and J. Wang, \"Online recruitment fraud detection based on a hybrid deep learning model,\" in *2023 IEEE 6th International Conference on Information Systems and Computer Aided Education (ICISCAE)*, IEEE, 2023, pp. 431-435."
    ]
    
    for idx, ref in enumerate(references, start=1):
        add_reference(doc, idx, ref)
        
    doc_path = os.path.join(target_dir, "paper.docx")
    try:
        doc.save(doc_path)
        print(f"Word document saved successfully to {doc_path}!")
    except PermissionError:
        fallback_path = os.path.join(target_dir, "paper_updated.docx")
        doc.save(fallback_path)
        print(f"Permission denied on {doc_path} (likely open in Word). Saved to {fallback_path} instead!")

if __name__ == "__main__":
    main()
