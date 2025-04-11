import requests
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
import re

# API credentials and configurations
api_token = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjQ0Mjk5MTA4NCwiYWFpIjoxMSwidWlkIjozNzM0Mzc2NSwiaWFkIjoiMjAyNC0xMi0wMVQxNTowNDo0MS44NDhaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6NjMyNDEyNCwicmduIjoidXNlMSJ9.zaheuo0ErsW3lC4bPF_am92YdjZeByTUNWJnhsk1lR8"
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
    "{{CompanyVATNumber}}": None,
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


def get_linked_item_details(linked_item_id):
    """ Helper function to get more details about a linked item """
    url = "https://api.monday.com/v2"

    query = '''
    query {{
        items(ids: [{linked_item_id}]) {{
            id
            name
            column_values {{
                text
                value
            }}
        }}
    }}
    '''.format(linked_item_id=linked_item_id)

    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json={"query": query})
    linked_item_data = response.json()

    return linked_item_data


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
                            ... on BoardRelationValue {{
                                linked_items {{
                                    id
                                    name
                                }}
                            }}
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

    # Extract items
    items = monday_data['data']["boards"][0]["groups"][0]['items_page']['items']

    for item in items:
        item_details = {
            'id': item['id'],
            'name': item['name'],
            'columns': [],
        }

        # Loop through each column to find linked items
        for column in item['column_values']:
            column_info = {
                'column_id': column.get('id', 'N/A'),  # Default to 'N/A' if id is not present
                'column_name': column.get('text', 'N/A'),  # Default to 'N/A' if text is not present
                'linked_items': []
            }

            # Check if the column has linked items
            if 'linked_items' in column:
                for linked_item in column['linked_items']:
                    # Fetch additional details about each linked item
                    linked_item_details = get_linked_item_details(linked_item['id'])

                    column_info['linked_items'].append({
                        'linked_item_id': linked_item['id'],
                        'linked_item_name': linked_item['name'],
                        'linked_item_details': linked_item_details['data']['items'][0]
                        # Additional details from linked item
                    })

            item_details['columns'].append(column_info)

        items_list.append(item_details)

    return items_list


def process_data_to_requested_structure(items_list):
    """Process the data into the requested structure with specific column mapping"""
    result = []

    for item in items_list:
        # Initialize the structured item
        structured_item = {
            'Item ID': item['id'],
            'Name': item['name'],
            'Columns': column_mapping.copy()  # Start with all None values
        }

        # Set the company name to the item name
        structured_item['Columns']['{{CompanyFullName}}'] = item['name']

        # Initialize counters to track which columns we've processed
        linked_item_count = 0  # Track position of linked items
        sales_manager_found = False
        ce_found = False
        templates_found = False
        action_found = False

        # Process each column in order
        column_position = 0
        mapping_keys = list(column_mapping.keys())

        # First pass - handle the simple text columns
        for i, column in enumerate(item['columns']):
            if column['column_name'] == item['name']:
                continue

            # For regular columns with text values that aren't relationships
            if not column['linked_items'] and column['column_name'] not in [None, 'N/A']:
                if i < len(mapping_keys) and structured_item['Columns'][mapping_keys[i]] is None:
                    structured_item['Columns'][mapping_keys[i]] = column['column_name']

        # Second pass - look for "Action" status column
        for column in item['columns']:
            if column['column_name'] and 'Generate' in column['column_name']:
                structured_item['Columns']['Action'] = column['column_name']
                action_found = True
                break

        # Third pass - handle linked items and specific columns
        for column in item['columns']:
            # Look for templates - they'll have multiple linked items
            if column['linked_items'] and len(column['linked_items']) > 1:
                urls = []
                for linked_item in column['linked_items']:
                    for column_value in linked_item['linked_item_details']['column_values']:
                        if column_value.get('text') and column_value['text'] is not None and 'https://' in str(
                                column_value['text']):
                            urls.append(column_value['text'])
                            break

                structured_item['Columns']['Contracts Templates'] = urls
                templates_found = True

            # Look for destination folder
            elif column['column_name'] and 'drive.google.com' in column['column_name']:
                structured_item['Columns']['Destination Folder'] = column['column_name']

            # Look for people (Sales Manager and CE)
            elif column['linked_items'] and len(column['linked_items']) == 1:
                linked_person = column['linked_items'][0]

                # If we haven't found the Sales Manager yet
                if not sales_manager_found:
                    structured_item['Columns']['{{SalesManagerName}}'] = linked_person['linked_item_name']

                    # Find email in linked_item_details
                    for column_value in linked_person['linked_item_details']['column_values']:
                        if column_value.get('text') and column_value['text'] and '@' in str(column_value['text']):
                            structured_item['Columns']['{{SalesManagerEmail}}'] = column_value['text']
                            break

                    sales_manager_found = True

                # If we've found Sales Manager but not CE yet
                elif not ce_found:
                    structured_item['Columns']['{{CEName}}'] = linked_person['linked_item_name']

                    # Find email in linked_item_details
                    for column_value in linked_person['linked_item_details']['column_values']:
                        if column_value.get('text') and column_value['text'] and '@' in str(column_value['text']):
                            structured_item['Columns']['{{CEEmail}}'] = column_value['text']
                            break

                    ce_found = True

        # Final pass - fill in the remaining columns based on position
        for i, column in enumerate(item['columns']):
            if i < len(mapping_keys) and column['column_name'] not in [None, 'N/A', item['name']]:
                key = mapping_keys[i]
                # Don't overwrite already set values
                if structured_item['Columns'][key] is None:
                    structured_item['Columns'][key] = column['column_name']

        result.append(structured_item)

    return result


def modify_google_doc(doc_id, replacements):
    requests = []
    for placeholder, value in replacements.items():
        if value is None:
            value = ""  # Handle None values
        elif isinstance(value, list):
            value = ", ".join(str(item) for item in value)  # Join lists into strings
        else:
            value = str(value)  # Convert any other types to string

        # Skip empty placeholders to avoid issues
        if not placeholder.strip():
            continue

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
        try:
            docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
            print(f"Successfully updated document {doc_id}")
        except Exception as e:
            print(f"Error updating document {doc_id}: {e}")
            # Print more details for debugging
            print(f"Request structure: {json.dumps(requests[:3], indent=2)}")  # Print first 3 for clarity



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


def send_completion_update(item_id, document_links):
    links_message = "\n".join([f"Link {index + 1}: {link}" for index, link in enumerate(document_links)])
    message = f"The task has been completed successfully. Here are the links to the new documents:\n{links_message}"

    # Escape double quotes and newlines for GraphQL
    escaped_message = message.replace('"', '\\"').replace('\n', ' ')

    url = "https://api.monday.com/v2"
    query = '''
    mutation {{
        create_update(item_id: {item_id}, body: "{message}") {{
            id
        }}
    }}
    '''.format(item_id=item_id, message=escaped_message)

    headers = {
        "Authorization": api_token,
        "Content-Type": "application/json"
    }
    requests.post(url, headers=headers, json={"query": query})

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


def main():
    items_list = get_items_from_board()
    output= process_data_to_requested_structure(items_list)
    print(output)
    for item in output:
        item_columns = item['Columns']
        update_item_status(item['Item ID'], "Working on it")
        replacements = {key: value for key, value in item_columns.items() if value}
        print(replacements)

        # Get the list of template URLs
        contracts_templates = item_columns.get("Contracts Templates", "")
        print(contracts_templates)
        destination_folder_url = item_columns.get("Destination Folder", "")

        folder_id_match = re.search(r'folders/([a-zA-Z0-9-_]+)|id=([a-zA-Z0-9-_]+)', destination_folder_url)
        destination_folder_id = folder_id_match.group(1) if folder_id_match else None

        template_urls = re.findall(r'(https?://[^\s),]+)', ' '.join(contracts_templates))


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
