require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md
session = GoogleDrive::Session.from_config("config.json")
# create folder
folder  = session.root_collection.create_subcollection("Uploaded files")
p folder
# Gets list of remote files.
session.files.each do |file|
  p file.title
end
# Uploads a local file.
# session.upload_from_file("/path/to/hello.txt", "hello.txt", convert: false)
# folder_id = folder.id 
# session.upload_from_file("/home/yuva/Downloads/InvoiceTemplate.docx", "hello.txt", convert: false)


# Downloads to a local file.
# name of drive file which you want to download
# file = session.file_by_title("hello.txt")
# path where you want to download this file 
# file.download_to_file("/home/yuva/Downloads/Neeraj/file/hello.txt")

# Updates content of the remote file.
# file.update_from_file("/path/to/hello.txt")


# upload a file inside specific folder 
# session.folder_by_url("https://drive.google.com/drive/folders/1akTg67Rl1CHLQUVOQnWwbrjwOa2lR4ny?ogsrc=32").upload_from_file("/home/yuva/Downloads/InvoiceTemplate.docx", "hello.txt", convert: false)



