import pickle
import os.path
from googleapiclient import errors
import email
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import re
import nltk
from nltk.corpus import stopwords
import os

stop = stopwords.words('english')

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

#change the path to the folder where credentials.json file is stored
os.chdir("D:/Work/Out_of_office/")
def search_message(service, user_id, search_string):
    """
    Search the inbox for emails using standard gmail search parameters
    and return a list of email IDs for each result
    PARAMS:
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        search_string: search operators you can use with Gmail
        (see https://support.google.com/mail/answer/7190?hl=en for a list)
    RETURNS:
        List containing email IDs of search query
    """
    try:
        # initiate the list for returning
        list_ids = []

        # get the id of all messages that are in the search string
        search_ids = service.users().messages().list(userId=user_id, q=search_string).execute()
        
        # if there were no results, print warning and return empty string
        try:
            ids = search_ids['messages']

        except KeyError:
            print("WARNING: the search queried returned 0 results")
            print("returning an empty string")
            return ""

        if len(ids)>1:
            for msg_id in ids:
                list_ids.append(msg_id['id'])
            return(list_ids)

        else:
            list_ids.append(ids['id'])
            return list_ids
        
    except (errors.HttpError, error):
        print("An error occured: %s")  % error


def get_message(service, user_id, msg_id):
    """
    Search the inbox for specific message by ID and return it back as a 
    clean string. String may contain Python escape characters for newline
    and return line. 
    
    PARAMS
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        msg_id: the unique id of the email you need
    RETURNS
        A string of encoded text containing the message body
    """
    try:
        # grab the message instance
        message = service.users().messages().get(userId=user_id, id=msg_id,format='raw').execute()

        # decode the raw string, ASCII works pretty well here
        msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))

        # grab the string from the byte object
        mime_msg = email.message_from_bytes(msg_str)
        
        from_= email.utils.parseaddr(mime_msg['From'])
        # check if the content is multipart (it usually is)
        content_type = mime_msg.get_content_maintype()
        if content_type == 'multipart':
            # there will usually be 2 parts the first will be the body in text
            # the second will be the text in html
            parts = mime_msg.get_payload()
            # from_1= email.utils.parseaddr(mime_msg['From'])
            # return the encoded text
            final_content = parts[0].get_payload()
            return final_content
        
        elif content_type == 'text':
            return mime_msg.get_payload()

        else:
            return ""
            print("\nMessage is not text or multipart, returned an empty string")
    # unsure why the usual exception doesn't work in this case, but 
    # having a standard Exception seems to do the trick
    except (errors.HttpError, error):
        print('An error occured: %s') % error 
        
    return from_, final_content


def get_service():
    """
    Authenticate the google api client and return the service object 
    to make further calls
    PARAMS
        None
    RETURNS
        service api object from gmail for making calls
    """
    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)


    service = build('gmail', 'v1', credentials=creds)

    return service

def extract_phone_numbers(string):
	"""
    Extract phone numbers from the email body
    PARAMS:
        string: body of the email
    RETURNS:
        List of phone numbers in the body
    """
    r= re.compile(r'\(?\b[2-9][0-9]{2}\)?[-. ]?[2-9][0-9]{2}[-. ]?[0-9]{4}\b')
    phone_numbers = r.findall(string)
    return [re.sub(r'\D', '', number) for number in phone_numbers]

def extract_email_addresses(string):
	"""
    Extract email addresses from the email body
    PARAMS:
        string: body of the email
    RETURNS:
        List of email addresses in the body
    """
    r = re.compile(r'[\w\.-]+@[\w\.-]+')
    return r.findall(string)

def ie_preprocess(string):
	"""
    POS tag the words in the email body and return a tagged list
    PARAMS:
        string: body of the email
    RETURNS:
        List of words with their POS tags
    """
    string = ' '.join([i for i in string.split() if i not in stop])
    sentences = nltk.sent_tokenize(string)
    sentences = [nltk.word_tokenize(sent) for sent in sentences]
    sentences = [nltk.pos_tag(sent) for sent in sentences]
    return sentences

def extract_names(string):
	"""
    Extract names from the email body
    PARAMS:
        string: body of the email
    RETURNS:
        List of names in the body
    """
    names = []
    sentences = ie_preprocess(string)
    for tagged_sentence in sentences:
        for chunk in nltk.ne_chunk(tagged_sentence):
            if type(chunk) == nltk.tree.Tree:
                if chunk.label() == 'PERSON':
                    names.append(' '.join([c[0] for c in chunk]))
    return names



 if __name__ == '__main__':
        service= get_service()
        list_ids= search_message(service, 'me', 'Automatic reply')
        final_content= [get_message(service, 'me', ids) for ids in list_ids]
        #take a sample email body, in this case 12th email in the list
        print("Number =", extract_phone_numbers(final_content[11].rstrip('\r'))) 
        print("Emails =", extract_email_addresses(final_content[11].rstrip('\r')))
        print("Names =", extract_names(final_content[11].rstrip('\r')))   