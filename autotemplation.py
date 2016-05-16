#!/usr/bin/env python3

import configparser
import httplib2
import io
import os
import re
import sys
from datetime import datetime

import oauth2client
from apiclient import discovery, errors
from apiclient.http import MediaFileUpload, MediaIoBaseDownload
from docxtpl import DocxTemplate
from jinja2 import Environment
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'autotemplation'
TEMPLATE_REGEX = r"\{\{ ([A-Za-z0-9_]+) \}\}"


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to {}'.format(credential_path))
    return credentials


def get_or_create_destination_folder_id(drive_service,
                                        destination_folder_name):
    folder_mime_type = 'application/vnd.google-apps.folder'
    response = drive_service.files().list(
        q="mimeType='{}' and name='{}'".format(folder_mime_type,
                                               destination_folder_name),
        spaces='drive',
        fields='nextPageToken, files(id, name)').execute()
    files = response.get('files', [])
    if files:
        folder_id = files[0].get('id')
    else:
        file_metadata = {
            'name': destination_folder_name, 'mimeType': folder_mime_type
        }
        file = drive_service.files().create(body=file_metadata,
                                            fields='id').execute()
        folder_id = file.get('id')
    return folder_id


def get_files_in_folder(service, folder_id):
    """Print files belonging to a folder.

    Args:
      service: Drive API service instance.
      folder_id: ID of the folder to print files from.
    """
    files = dict()
    page_token = None
    while True:
        try:
            param = {}
            if page_token:
                param['pageToken'] = page_token
            children = service.files().list(
                q="'{}' in parents and trashed=false".format(folder_id),
                spaces='drive',
                fields='nextPageToken, files(id, name)',
                pageToken=page_token).execute()
            for child in children.get('files', []):
                files[child['name']] = child['id']
            page_token = children.get('nextPageToken')
            if not page_token:
                break
        except errors.HttpError as error:
            print('An error occurred: {}'.format(error))
            break
    return files


def get_template(service, folder_id):
    templates = get_files_in_folder(service, folder_id)
    if len(templates) == 0:
        print("No Templates Found. Please check Folder ID. Exiting.")
        sys.exit(1)
    elif len(templates) == 1:
        first = next(iter(templates.items()))
        print("One Template Found. Using {}.".format(first[0]))
        return first
    template_list = sorted(templates.keys())
    print('Please Select a Template:')
    while True:
        for i, name in enumerate(template_list):
            print('{}) {}'.format(i+1, name))
        print('q) Quit')
        choice = input(
            'Selection?  '.format(len(template_list)))
        if choice.lower() in ['q', 'quit', 'exit']:
            print('Exiting.')
            sys.exit(0)
        try:
            file_name = template_list[int(choice)-1]
            file_id = templates[file_name]
            break
        except (ValueError, IndexError):
            print('Error: Invalid Selection.')
    return file_name, file_id


def get_date_and_set_context(context_dict):
    use_date = input("Please enter document date or leave empty for today's "
                     "date (Ex: 20150120):  ")
    date_object = None
    if use_date:
        try:
            date_object = datetime.strptime(use_date, '%Y%m%d')
        except AttributeError:
            pass
    if not date_object:
        date_object = datetime.now()
    context_dict['DATE_FULL'] = date_object.strftime('%B %d, %Y')
    context_dict['DATE_FULL_NUM'] = date_object.strftime('%Y%m%d')
    context_dict['DATE_FULL_DASH'] = date_object.strftime('%m-%d-%Y')
    context_dict['DATE_FULL_SLASH'] = date_object.strftime('%m/%d/%Y')
    context_dict['DATE_MONTH'] = date_object.strftime('%B')
    context_dict['DATE_MONTH_NUM'] = date_object.strftime('%m')
    context_dict['DATE_DAY_FULL'] = date_object.strftime('%A')
    context_dict['DATE_DAY_SHORT'] = date_object.strftime('%a')
    context_dict['DATE_DAY_NUM'] = date_object.strftime('%d')
    context_dict['DATE_YEAR'] = date_object.strftime('%Y')


def get_target_name(template_name, context):
    return Environment().from_string(template_name).render(context)


def get_template_variables(doc, template_name):
    template_vars = set()

    #  Template Name
    template_vars |= set(re.findall(TEMPLATE_REGEX, template_name))
    #  Paragraphs
    template_vars |= set([var for var_list in
                          [re.findall(TEMPLATE_REGEX, p.text)
                           for p in doc.paragraphs] for var in var_list])
    #  Tables
    template_vars |= set([var for var_list in
                          [re.findall(TEMPLATE_REGEX, cell.text)
                           for table in doc.tables
                           for row in table.rows
                           for cell in row.cells] for var in var_list])
    return template_vars


def main():
    config = configparser.ConfigParser()
    config.read('autotemplation.ini')
    template_folder_id = config['DEFAULT']['TemplateFolderID']
    destination_folder_name = config['DEFAULT']['DestinationFolderName']
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    drive_service = discovery.build('drive', 'v3', http=http)
    word_type = 'vnd.openxmlformats-officedocument.wordprocessingml.document'
    mime_type = 'application/{}'.format(word_type)
    destination_folder_id = get_or_create_destination_folder_id(
        drive_service, destination_folder_name)
    template_file_name, template_file_id = get_template(drive_service,
                                                        template_folder_id)
    request = drive_service.files().export_media(
        fileId=template_file_id,
        mimeType=mime_type)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    doc = DocxTemplate(fh)
    full_doc = doc.get_docx()
    template_vars = get_template_variables(full_doc, template_file_name)
    context = dict()
    get_date_and_set_context(context)
    for var in template_vars:
        if var not in context:
            context[var] = input("Enter a value for {}:  ".format(var))
    new_file_name = get_target_name(template_file_name, context)
    doc.render(context)
    docx_name = '{}.docx'.format(new_file_name)
    pdf_name = '{}.pdf'.format(new_file_name)
    doc.save(docx_name)
    file_metadata = {
        'name': new_file_name,
        'parents': [destination_folder_id],
        'mimeType': 'application/vnd.google-apps.document'
    }
    media = MediaFileUpload(docx_name,
                            mimetype=mime_type,
                            resumable=True)
    file = drive_service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()
    print('{} placed in folder {}.'.format(new_file_name,
                                           destination_folder_name))


if __name__ == '__main__':
    main()
