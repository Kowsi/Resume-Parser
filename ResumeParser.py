import docx2txt
import spacy
import re
from spacy.matcher import Matcher
from spacy.tokens import Span

class MatchEvent:
    PERSON_PATTERN      = [{'POS': 'PROPN', 'ENT_TYPE': 'PERSON'},
                          {'POS': 'PROPN', 'ENT_TYPE': 'PERSON', 'OP': '?'},
                          {'POS': 'PROPN', 'ENT_TYPE': 'PERSON'}]

    EMAIL_ID_PATTERN    = [{"LIKE_EMAIL": True}]

    PHONE_PATTERN_1     = [{"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                           {"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                           {"SHAPE": "dddd"}]

    PHONE_PATTERN_2     = [{"ORTH": "("}, {"SHAPE": "ddd"}, {"ORTH": ")"},
                           {"SHAPE": "ddd"}, {"ORTH": "-", "OP": "?"},
                           {"SHAPE": "dddd", "LENGTH": {"==": 4}}]

    ADDRESS_PATTERN     = [{'POS': 'NUM'}, {'POS': 'PROPN'},
                           {'POS': 'PROPN', 'OP': "*"},
                           {'IS_PUNCT': True, 'OP': '?'},
                           {'LEMMA': {'IN': ['apt', 'unit']}, 'OP': '?'},
                           {'POS': 'NUM'},
                           {'ORTH': ',', 'OP': "?"},
                           {'POS': 'PROPN'},
                           {'ORTH': {
                               'REGEX': '(AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VI|VT|WA|WI|WV|WY)'}}]

    LINKEDIN_URL_PATTERN = [{"LIKE_URL": True, 'ORTH': {'REGEX': r'\s*(linkedin.com)\s*'}}]

    GIT_URL_PATTERN     = [{"LIKE_URL": True, 'ORTH': {'REGEX': r'\s*(github.com)\s*'}}]

    @classmethod
    def full_name_event(cls, matcher, doc, i, matches):
        match_id, start, end = matches[i]
        entity = Span(doc, start, end, label="EVENT")
        if i == 2:
            full_name = doc[matches[0][1]:matches[0][2]]
            if str(entity.text).startswith(str(full_name)):
                matches[0] = matches[i]


class ResumeParser:

    nlp = spacy.load("en_core_web_sm")

    CANDIDATE_INFO = [{'id': 'FullName',    'match_on': MatchEvent.full_name_event, 'pattern': MatchEvent.PERSON_PATTERN},
                      {'id': 'Email',       'match_on': None,                       'pattern': MatchEvent.EMAIL_ID_PATTERN},
                      {'id': 'Address',     'match_on': None,                       'pattern': MatchEvent.ADDRESS_PATTERN},
                      {'id': 'Phone',       'match_on': None,                       'pattern': MatchEvent.PHONE_PATTERN_1},
                      {'id': 'Phone',       'match_on': None,                       'pattern': MatchEvent.PHONE_PATTERN_2},
                      {'id': 'GithubURL',   'match_on': None,                       'pattern': MatchEvent.GIT_URL_PATTERN},
                      {'id': 'LinkedInURL', 'match_on': None,                       'pattern': MatchEvent.LINKEDIN_URL_PATTERN}]

    def __init__(self, file_name):
        self.txt = self.convert_docx2txt(file_name)
        self.doc = ResumeParser.nlp(self.txt)
        self.matcher = Matcher(self.nlp.vocab, validate=True)


    # Converting Docx to txt using docx2txt
    def convert_docx2txt(self, file_name):
        temp = docx2txt.process(file_name)
        text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
        return ' '.join(text)

    def get_candidate_info(self):
        # To store Candidate information
        candidate_info_details = {}

        # Adding all the patterns to the Matcher to retrieve the corresponding details
        for index, info in enumerate(ResumeParser.CANDIDATE_INFO):
            self.matcher.add(info['id'], info['match_on'], info['pattern'])
            candidate_info_details[info['id']] = 'Null'

        # Extracting the details from document text
        matches = self.matcher(self.doc)

        # Iterating the result set and store the information
        for match_id, start, end in matches:
            rule_id = str(self.nlp.vocab.strings[match_id])
            if candidate_info_details[rule_id] == 'Null':
                span = self.doc[start:end]
                candidate_info_details[rule_id] = str(span.text)

        #for key in candidate_info_details:
        #   print(key + " : " + candidate_info_details[key])

        return {'CandidateInformation': candidate_info_details}



def main():
    rp = ResumeParser('./Sample Data/Christina_Austin_BA.docx')
    print(rp.get_candidate_info())
    rp = ResumeParser('./Sample Data/Deepak Kumar_Business Analyst_PM-NC.docx')
    print(rp.get_candidate_info())
    rp = ResumeParser('./Sample Data/Manish_ARORA_Profile.docx')
    print(rp.get_candidate_info())

main()
