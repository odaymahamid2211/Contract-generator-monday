import requests
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
import re

# API credentials and configurations
api_token = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjM5MzQ5Mzg3NiwiYWFpIjoxMSwidWlkIjo0MDMzODA5MCwiaWFkIjoiMjAyNC0wOC0wNlQxMTozMzowMC4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjMyNDEyNCwicmduIjoidXNlMSJ9.0u1Wk85mFM50YptDZ65-6maJl0Jwhlhtzb8rYXFs8I0"  # Replace with your actual token
board_id = "7167670399"
service_account_file = "C:\\Users\\oday\\Desktop\\Contracts\\pythonProject\\generate-reports.json"  # Path to your service account JSON
scopes = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
]

# Column mapping for item details
column_mapping = {
    "{{CompanyFullName}}": None,
    "{{CompanyDomain}}": None,
    "{{CompanyNumber}}": None,
    "{{CompanyLawsState}}": None,
    "{{CompanyVATNumber}}" : None,
    "{{CompanyAddress}}": None,
    "{{QuoteDate}}": None,
    "{{Currency}}": None,
    "{{PaymentDaysTerm}}": None,
    "{{ContactEmail}}": None,
    "{{SalesManagerName}}": None,
    "{{SalesManagerEmail}}": None,
    "{{CEName}}": None,
    "{{CEEmail}}": None,
    "Destination Folder": None,
    "Contracts Templates": None,
    "Action": None,
}

# Set up Google API client
credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
docs_service = build('docs', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)

def get_items_from_board():
    url = "https://api.monday.com/v2"
    query = '''
    query {{
        boards(ids: [{board_id}]) {{
            groups(ids: ["topics"]) {{
                id
                title
                items_page {{
                    items {{
                        id
                        name
                        column_values {{
                            text
                            value
                        }}
                    }}
                }}
            }}
        }}
    }}
    '''.format(board_id=board_id)

    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json={"query": query})
    monday_data = response.json()
    items_list = []

    items = monday_data['data']["boards"][0]["groups"][0]['items_page']['items']
    for item in items:
        item_id = item['id']
        item_name = item['name']
        column_values = []

        for column in item['column_values']:
            text_value = column.get('text', '')
            json_value = column.get('value')

            if json_value:
                parsed_value = json.loads(json_value)
                if 'linkedPulseIds' in parsed_value:
                    linked_items = parsed_value['linkedPulseIds']
                    linked_item_details = []
                    for linked_item in linked_items:
                        linked_item_id = linked_item['linkedPulseId']
                        linked_details = fetch_linked_item_details(linked_item_id)
                        linked_item_details.append(linked_details)
                    column_values.append(", ".join(linked_item_details))
                else:
                    column_values.append(text_value)
            else:
                column_values.append(text_value)

        item_details_dict = map_columns_to_dict(column_values)
        items_list.append({"Item ID": item_id, "Name": item_name, "Columns": item_details_dict})

    return items_list


def map_columns_to_dict(column_values):
    mapped_dict = column_mapping.copy()
    columns_list = list(mapped_dict.keys())

    for index, value in enumerate(column_values):
        if index < len(columns_list):
            mapped_dict[columns_list[index]] = value

    # Extract SalesManagerName and SalesManagerEmail from the formatted text
    sales_manager_details = mapped_dict.get("{{SalesManagerName}}", "")
    if sales_manager_details:
        match = re.match(r'([^(]+)\s*\(.*?,\s*([^,]+)', sales_manager_details)
        if match:
            mapped_dict["{{SalesManagerName}}"] = match.group(1).strip()
            mapped_dict["{{SalesManagerEmail}}"] = match.group(2).strip()

    # Extract CEName and CEmail from the formatted text
    ce_details = mapped_dict.get("{{CEName}}", "")
    if ce_details:
        match = re.match(r'([^(]+)\s*\(.*?,\s*([^,]+)', ce_details)
        if match:
            mapped_dict["{{CEName}}"] = match.group(1).strip()
            mapped_dict["{{CEEmail}}"] = match.group(2).strip()

    return mapped_dict



def fetch_linked_item_details(item_id):
    url = "https://api.monday.com/v2"
    query = '''
    query {{
        items(ids: {item_id}) {{
            id
            name
            column_values {{
                text
            }}
        }}
    }}
    '''.format(item_id=item_id)

    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json={"query": query})
    linked_item_data = response.json()

    if 'data' in linked_item_data and linked_item_data['data']['items']:
        item_details = linked_item_data['data']['items'][0]
        linked_item_name = item_details['name']
        links = [col['text'] for col in item_details['column_values'] if col['text']]
        return f"{linked_item_name} ({', '.join(links)})"

    return f"Linked Item ID: {item_id}"

def modify_google_doc(doc_id, replacements):
    requests = []
    for placeholder, value in replacements.items():
        requests.append({
            'replaceAllText': {
                'containsText': {
                    'text': placeholder,
                    'matchCase': True
                },
                'replaceText': value
            }
        })

    if requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()


def copy_and_modify_template(template_url, destination_folder_id, replacements, item_name, item_id):
    doc_id = template_url.split('/d/')[1].split('/')[0]
    print(f"Copying document with ID: {doc_id}")

    try:
        document = drive_service.files().get(fileId=doc_id, supportsAllDrives=True).execute()

        new_file_name = f"{item_name}-{document['name']}"

        copied_file = drive_service.files().copy(
            fileId=doc_id,
            body={
                'name': new_file_name,
                'parents': [destination_folder_id]
            },
            supportsAllDrives=True
        ).execute()

        print(f"Successfully copied to '{copied_file['name']}' with ID: {copied_file['id']}")

        modify_google_doc(copied_file['id'], replacements)

        return f"https://docs.google.com/document/d/{copied_file['id']}/edit"

    except Exception as e:
        print(f"Error copying or modifying document ID {doc_id}: {e}")
        update_item_status(item_id, "Stuck")  # Update status to "Stuck"
        return None


def update_item_status(item_id, status_label):
    url = "https://api.monday.com/v2"
    query = '''
    mutation {{
        change_column_value(board_id: {board_id}, item_id: {item_id}, column_id: "status__1", value: "{{\\"label\\": \\"{status_label}\\"}}") {{
            id
        }}
    }}
    '''.format(board_id=board_id, item_id=item_id, status_label=status_label)

    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json"
    }
    requests.post(url, headers=headers, json={"query": query})


def send_completion_update(item_id, document_links):
    links_message = "\n".join([f"Link {index + 1}: {link}" for index, link in enumerate(document_links)])
    message = f"The task has been completed successfully. Here are the links to the new documents:\n{links_message}"

    url = "https://api.monday.com/v2"
    query = '''
    mutation {{
        create_update(item_id: {item_id}, body: "{message}") {{
            id
        }}
    }}
    '''.format(item_id=item_id, message=message.replace('"', '\\"'))

    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json"
    }
    requests.post(url, headers=headers, json={"query": query})


def main():
    output = get_items_from_board()
    print(output)
    for item in output:
        item_columns = item['Columns']
        update_item_status(item['Item ID'], "Working on it")

        replacements = {key: value for key, value in item_columns.items() if value}

            # Get the list of template URLs
        contracts_templates = item_columns.get("Contracts Templates", "")
        destination_folder_url = item_columns.get("Destination Folder", "")

        folder_id_match = re.search(r'folders/([a-zA-Z0-9-_]+)|id=([a-zA-Z0-9-_]+)', destination_folder_url)
        destination_folder_id = folder_id_match.group(1) if folder_id_match else None

        template_urls = re.findall(r'(https?://[^\s),]+)', contracts_templates)

        document_links = []
        for template_url in template_urls:
            if destination_folder_id:
                 doc_link = copy_and_modify_template(
                       template_url=template_url,
                       destination_folder_id=destination_folder_id,
                       replacements=replacements,
                       item_name=item['Name'],
                       item_id=item['Item ID']
                    )
                 if doc_link:
                        document_links.append(doc_link)

        send_completion_update(item['Item ID'], document_links)
        update_item_status(item['Item ID'], "Generate completed")

if __name__ == "__main__":
    main()
