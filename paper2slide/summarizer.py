from transformers import *
from summarizer import Summarizer
from .preprocessor import get_cleaned_text, replace_et_al, split_sentences

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


