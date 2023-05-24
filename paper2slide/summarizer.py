from transformers import *
from summarizer import Summarizer
from .preprocessor import get_cleaned_text, replace_et_al

def summarize_text(output):
    # Load model, model config and tokenizer via Transformers
    custom_config = AutoConfig.from_pretrained('allenai/scibert_scivocab_cased')
    custom_config.output_hidden_states=True
    custom_tokenizer = AutoTokenizer.from_pretrained('allenai/scibert_scivocab_cased')
    custom_model = AutoModel.from_pretrained('allenai/scibert_scivocab_cased', config=custom_config)

    model = Summarizer(custom_model=custom_model, custom_tokenizer=custom_tokenizer)
    for i in range(len(output)):
        output[i]["content"] = get_cleaned_text(output[i]["content"])
        # et al. 때문에 .으로 문장을 구분하는 방식에 어려움 존재 -> 먼저 제거 후 다시 삽입

    for i in range(len(output)):
        if output[i]["role"] is None:
            res = model.calculate_optimal_k(output[i]["content"], k_max=20)
            output[i]["summarized"] = replace_et_al(model(output[i]["content"],num_sentences=res))
    
    return output


