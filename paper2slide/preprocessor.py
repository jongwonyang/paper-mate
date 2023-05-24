import re
from gingerit.gingerit import GingerIt

def get_cleaned_text(text):
    text = remove_et_al_period(text)
    text = remove_hyphen_spaces(text)
    text = correct_grammar(text)
    return text

def remove_et_al_period(text):
    # "et al."이 포함된 문장에서 온점을 찾아 제거합니다.
    text = re.sub(r'(?<=et al)\.(?=\s|$)', '', text)
    text = re.sub(r'\([^()]*\)', '', text)
    return text

def remove_hyphen_spaces(text):
    # -로 이어진 단어들의 패턴을 정의
    pattern = re.compile(r'(\w+)-\s+(\w+)')
    # 패턴과 일치하는 문자열을 검색하여 공백 제거
    corrected_text = pattern.sub(r'\1\2', text)
    # 결과 반환
    return corrected_text

def replace_et_al(text):
    # "et al" 패턴을 정의
    pattern = re.compile(r'et al(?![.\w])', re.IGNORECASE)
    # 패턴과 일치하는 문자열을 검색하여 "et al."로 대체
    corrected_text = pattern.sub('et al.', text)
    # 결과 반환
    return corrected_text

def correct_grammar(text):
    # 문장을 온점 기준으로 분리
    sentences = text.split('.')
    # 문장의 마지막이 온점이 아니라면, 마지막 문장에서 온점을 추가
    if sentences[-1] != '':
        sentences[-1] += '.'
    # GingerIt 클래스의 인스턴스 생성
    parser = GingerIt()
    # 수정된 문장을 저장할 리스트
    corrected_sentences = []
    # 각 문장에 대해 문법 검사 수행
    for sentence in sentences:
        # 공백 제거
        sentence = sentence.strip()
        # 빈 문자열이라면 다음 문장으로 건너뜀
        if sentence == '':
            continue
        # 문법 오류 수정
        result = parser.parse(sentence)
        corrected_sentence = result['result']
        corrected_sentences.append(corrected_sentence)
    # 수정된 문장을 온점으로 이어붙임
    result_text = '. '.join(corrected_sentences)
    # 결과 반환
    return result_text