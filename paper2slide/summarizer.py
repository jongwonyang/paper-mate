from transformers import *
from summarizer import Summarizer
from .preprocessor import get_cleaned_text, replace_et_al, split_sentences
from keybert import KeyBERT

def summarize_text(output):
    # Load model, model config and tokenizer via Transformers
    custom_config = AutoConfig.from_pretrained('allenai/scibert_scivocab_cased')
    custom_config.output_hidden_states=True
    custom_tokenizer = AutoTokenizer.from_pretrained('allenai/scibert_scivocab_cased')
    custom_model = AutoModel.from_pretrained('allenai/scibert_scivocab_cased', config=custom_config)

    model = Summarizer(custom_model=custom_model, custom_tokenizer=custom_tokenizer)


    for i in range(len(output)):
        if output[i]["role"] is None:
            res = model.calculate_optimal_k(output[i]["content"], k_max=20)
            summarized = model(output[i]["content"],num_sentences=res)
            summarized_list = split_sentences(summarized)
            for j in range(len(summarized_list)):
                summarized_list[j] = replace_et_al(summarized_list[j])
            output[i]["summarized"] = summarized_list
    
    return output

def extract_keywords_from_paragraph(paragraph):
    # Initialize the KeyBERT model
    model = KeyBERT("distilbert-base-nli-mean-tokens")

    # Extract keywords from the paragraph
    keywords = model.extract_keywords(paragraph, keyphrase_ngram_range=(1, 2), stop_words='english', use_mmr = True,top_n = 20,diversity = 0.3)

    # Return the extracted keywords
    return keywords

