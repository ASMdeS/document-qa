import streamlit as st
import google.generativeai as genai
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path

# --- Configuration ---
# Define the scopes required by the application.
# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/drive.file', # To create and manage files (including sharing)
    'https://www.googleapis.com/auth/documents' # To create and edit Google Docs
]
CREDENTIALS_FILE = 'credentials.json' # Download from Google Cloud Console
TOKEN_PICKLE_FILE = 'token.pickle'

# --- Helper Functions ---

def authenticate_google():
    """Handles Google OAuth2 authentication flow."""
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, 'rb') as token:
            try:
                creds = pickle.load(token)
            except EOFError:
                st.error("Error loading saved credentials. Please re-authenticate.")
                creds = None # Force re-authentication
            except Exception as e:
                 st.error(f"An unexpected error occurred loading credentials: {e}")
                 creds = None # Force re-authentication


    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                st.error(f"Failed to refresh token: {e}. Please re-authenticate.")
                # Attempt the full flow if refresh fails
                if os.path.exists(CREDENTIALS_FILE):
                    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                    # Note: run_local_server() opens a browser tab for auth.
                    # This works well locally but needs adjustment for deployed Streamlit apps (e.g., using st.experimental_get_query_params for callback).
                    creds = flow.run_local_server(port=0)
                else:
                    st.error("Credentials file (credentials.json) not found.")
                    return None
        else:
            if os.path.exists(CREDENTIALS_FILE):
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                # For Streamlit Cloud or other deployments, you'd need a different flow
                # that handles redirects properly without a local server.
                # This might involve manual URL copy/paste or more complex setup.
                st.info("Please follow the authentication steps in your browser.")
                creds = flow.run_local_server(port=0)
            else:
                st.error("Credentials file (credentials.json) not found. Please download it from Google Cloud Console and place it in the app directory.")
                return None

        # Save the credentials for the next run
        with open(TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(creds, token)

    return creds

def create_google_doc(service, title):
    """Creates a blank Google Doc with the given title."""
    try:
        body = {'title': title}
        doc = service.documents().create(body=body).execute()
        st.success(f"Document '{doc.get('title')}' created successfully.")
        return doc.get('documentId')
    except HttpError as error:
        st.error(f"An error occurred creating the document: {error}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred during document creation: {e}")
        return None


def insert_text_into_doc(service, document_id, text):
    """Inserts text into the beginning of the Google Doc."""
    try:
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': 1,  # Start at the beginning of the document body
                    },
                    'text': text
                }
            }
        ]
        service.documents().batchUpdate(
            documentId=document_id, body={'requests': requests}).execute()
        st.success("Curriculum content inserted into the document.")
        return True
    except HttpError as error:
        st.error(f"An error occurred inserting text: {error}")
        return False
    except Exception as e:
        st.error(f"An unexpected error occurred during text insertion: {e}")
        return False

def share_google_doc(service, file_id):
    """Shares the Google Doc so anyone with the link can view."""
    try:
        permission = {
            'type': 'anyone',
            'role': 'reader' # Or 'writer' if you want them to edit
        }
        service.permissions().create(
            fileId=file_id,
            body=permission,
            fields='id' # We just need to know if it succeeded
        ).execute()

        # Get the web link for sharing
        file_metadata = service.files().get(fileId=file_id, fields='webViewLink').execute()
        share_link = file_metadata.get('webViewLink')
        st.success("Document sharing permissions updated.")
        return share_link
    except HttpError as error:
        st.error(f"An error occurred sharing the document: {error}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred during document sharing: {e}")
        return None

# --- Streamlit App UI ---

st.set_page_config(page_title="Curriculum Generator", layout="wide")
st.title("📝 AI Curriculum Generator")
st.write("Enter the details below, and the AI will generate a curriculum, save it as a Google Doc, and provide a shareable link.")
st.markdown("---")

# --- API Key Input ---
st.sidebar.header("API Keys")
# It's recommended to use st.secrets for API keys in deployed apps
# For local testing, text input is used here.
gemini_api_key = st.sidebar.text_input("Enter your Google Gemini API Key:", type="password")

# --- Curriculum Input ---
st.header("Curriculum Details")
subject = st.text_input("Subject / Course Title:", "Introduction to Python Programming")
target_audience = st.text_input("Target Audience:", "Beginners with no prior programming experience")
duration = st.text_input("Course Duration (e.g., 8 weeks, 1 semester):", "8 weeks")
learning_objectives = st.text_area("Key Learning Objectives (one per line):", """Understand fundamental programming concepts (variables, data types, control flow)
Write basic Python scripts
Work with common data structures (lists, dictionaries)
Read from and write to files
Understand the basics of functions""")
key_topics = st.text_area("Key Topics/Modules (one per line):", """Week 1: Introduction, Setup, Variables, Basic Data Types
Week 2: Operators, Expressions, Input/Output
Week 3: Control Flow (if/else, loops)
Week 4: Lists and Tuples
Week 5: Dictionaries and Sets
Week 6: Functions
Week 7: File Handling
Week 8: Basic Error Handling, Course Wrap-up""")
additional_instructions = st.text_area("Any other specific instructions for the AI?", "Include suggestions for simple weekly exercises or mini-projects.")

st.markdown("---")

# --- Generation Button and Output ---
if st.button("Generate Curriculum & Google Doc"):
    if not gemini_api_key:
        st.warning("Please enter your Gemini API Key in the sidebar.")
    else:
        try:
            # Configure Gemini API
            genai.configure(api_key=gemini_api_key)
            model = genai.GenerativeModel('gemini-pro') # Or the latest appropriate model

            # Authenticate Google Account
            with st.spinner("Authenticating with Google... Please check your browser."):
                creds = authenticate_google()

            if creds:
                st.success("Google Authentication Successful!")
                try:
                    # Build Google Services
                    docs_service = build('docs', 'v1', credentials=creds)
                    drive_service = build('drive', 'v3', credentials=creds)

                    # --- Generate Curriculum Content ---
                    with st.spinner("Generating curriculum content with Gemini AI..."):
                        prompt = f"""
                        Generate a detailed curriculum outline based on the following information:

                        Subject/Course Title: {subject}
                        Target Audience: {target_audience}
                        Course Duration: {duration}
                        Learning Objectives:
                        {learning_objectives}

                        Key Topics/Modules:
                        {key_topics}

                        Additional Instructions: {additional_instructions}

                        Please structure the output clearly, perhaps week by week or module by module, including topics covered, activities, and potential assessments for each section.
                        Ensure the output is well-formatted plain text suitable for a document.
                        """
                        response = model.generate_content(prompt)
                        generated_curriculum = response.text
                        st.success("Curriculum content generated.")
                        # st.subheader("Generated Curriculum Preview:")
                        # st.text_area("Preview", generated_curriculum, height=300) # Optional preview

                    # --- Create and Populate Google Doc ---
                    with st.spinner("Creating Google Doc..."):
                        doc_title = f"Curriculum: {subject}"
                        document_id = create_google_doc(docs_service, doc_title)

                    if document_id:
                        with st.spinner("Inserting content into Google Doc..."):
                            inserted = insert_text_into_doc(docs_service, document_id, generated_curriculum)

                        if inserted:
                            # --- Share Google Doc ---
                            with st.spinner("Sharing Google Doc..."):
                                share_link = share_google_doc(drive_service, document_id)

                            if share_link:
                                st.subheader("✅ Success!")
                                st.markdown(f"Your curriculum has been generated and saved as a Google Doc. You can access it here:")
                                st.markdown(f"**[{doc_title}]({share_link})**")
                                st.balloons()
                            else:
                                st.error("Failed to get the shareable link for the document.")
                        else:
                             st.error("Failed to insert content into the Google Doc.")
                    else:
                        st.error("Failed to create the Google Doc.")

                except HttpError as error:
                    st.error(f"An API error occurred: {error}")
                except Exception as e:
                    st.error(f"An unexpected error occurred: {e}")
            else:
                st.error("Google Authentication Failed. Cannot proceed.")
        except Exception as e:
            st.error(f"An error occurred configuring Gemini or during the generation process: {e}")

st.markdown("---")
st.caption("Ensure 'credentials.json' is in the correct directory and you have provided the Gemini API Key.")
