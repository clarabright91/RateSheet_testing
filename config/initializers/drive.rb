# require "google_drive"

# setup drive access
# drive = GoogleDrive::Session.from_config("config.json")
# check folder exists or not 
# folder_exist = false
# drive.folders.each do |folder|
# 	folder_exist = true if folder.title.eql?("Uploaded files")
# end

# create folder on drive
# unless folder_exist
# 	drive.root_collection.create_subcollection("Uploaded files")
# end
# require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md

# session = GoogleDrive::Session.from_config("config.json")
# create folder
# folder  = session.root_collection.create_subcollection("Uploaded files")
# p folder
# Gets list of remote files.
# session.files.each do |file|
#   p file.title
# end
# Uploads a local file.
# session.upload_from_file("/path/to/hello.txt", "hello.txt", convert: false)
# folder_id = folder.id 
# session.upload_from_file("/home/yuva/Downloads/OB_New_Penn_Financial_Wholesale5806.xls", "OB_New_Penn_Financial_Wholesale5806.xls", convert: false)
session = GoogleDrive::Session.from_config("config.json")
# create folder
# folder  = session.root_collection.create_subcollection("Uploaded files")
# p folder
# Gets list of remote files.
# session.files.each do |file|
#   p file.title
# end
# Uploads a local file.
# session.upload_from_file("/path/to/hello.txt", "hello.txt", convert: false)
# folder_id = folder.id 
# session.upload_from_file("/home/yuva/Downloads/OB_New_Penn_Financial_Wholesale5806.xls", "OB_New_Penn_Financial_Wholesale5806.xls", convert: false)

# Downloads to a local file.
# name of drive file which you want to download
# file = session.file_by_title("hello.txt")
# path where you want to download this file 
# file.download_to_file("/home/yuva/Downloads/Neeraj/file/hello.txt")

# Updates content of the remote file.
# file.update_from_file("/path/to/hello.txt")

# upload a file inside specific folder
# new_session = GoogleDrive::Session.from_config("config.json").folder_by_id("1p9gWPr4DPKgdQinWm7F8fU-hNVvKKaDM").upload_from_file("/home/yuva/Desktop/pureloan/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806.xls", "OB_New_Penn_Financial_Wholesale5806.xls", convert: false)
# session.folder_by_url("https://drive.google.com/drive/u/1/folders/1p9gWPr4DPKgdQinWm7F8fU-hNVvKKaDM?ogsrc=32").upload_from_file("/home/yuva/Downloads/InvoiceTemplate.docx", "hello.txt", convert: false)
# session.folder_by_url("https://drive.google.com/drive/u/1/folders/1p9gWPr4DPKgdQinWm7F8fU-hNVvKKaDM?ogsrc=32").upload_from_file("/home/yuva/Desktop/pureloan/RateSheetExtractor/OB_New_Penn_Financial_Wholesale5806.xls", "OB_New_Penn_Financial_Wholesale5806.xls", convert: false)

# controller code
# drive_access = GoogleDrive::Session.from_config("config.json")
# folder = drive_access.folder_by_id("1p9gWPr4DPKgdQinWm7F8fU-hNVvKKaDM")
# file = folder.file_by_title("OB_New_Penn_Financial_Wholesale5806.xls")

# new_session.spreadsheet_by_title("OB_New_Penn_Financial_Wholesale5806").web_view_link
# ws = drive_access.spreadsheet_by_key("1q8sGJ7h3x8w6F5knzLFiST_FzNDVWYYOvkNHrf7zN9M").worksheets[0]


# require 'open-uri'
# open('image.png', 'wb') do |file|
#   file << open('http://example.com/image.png').read
# end

# IO.copy_stream(open('http://example.com/image.png'), 'destination.png')

# require 'open-uri'
# download = open('http://example.com/image.png')

# IO.copy_stream(download, '~/image.png')

# controller code
# drive_access = GoogleDrive::Session.from_config("config.json")
# folder = drive_access.folder_by_id("1p9gWPr4DPKgdQinWm7F8fU-hNVvKKaDM")
# file = folder.file_by_title("OB_New_Penn_Financial_Wholesale5806.xls")



# new_session.spreadsheet_by_title("OB_New_Penn_Financial_Wholesale5806").web_view_link
# ws = drive_access.spreadsheet_by_key("1q8sGJ7h3x8w6F5knzLFiST_FzNDVWYYOvkNHrf7zN9M").worksheets[0]

# require 'open-uri'
# open('image.png', 'wb') do |file|
#   file << open('http://example.com/image.png').read
# end

# IO.copy_stream(open('http://example.com/image.png'), 'destination.png')

# require 'open-uri'
# download = open('http://example.com/image.png')
# IO.copy_stream(download, '~/image.png')


# controller code
# drive_access = GoogleDrive::Session.from_config("config.json")
# folder = drive_access.folder_by_id("1p9gWPr4DPKgdQinWm7F8fU-hNVvKKaDM")
# file = folder.file_by_title("OB_New_Penn_Financial_Wholesale5806.xls")



# new_session.spreadsheet_by_title("OB_New_Penn_Financial_Wholesale5806").web_view_link
# ws = drive_access.spreadsheet_by_key("1q8sGJ7h3x8w6F5knzLFiST_FzNDVWYYOvkNHrf7zN9M").worksheets[0]
