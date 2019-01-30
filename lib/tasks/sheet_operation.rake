require 'open-uri'
require "google_drive"
require 'roo'
require 'roo-xls'

namespace :sheet_operation do
	desc "export sheet and import sheet"
	#setup drive access
	drive  			 = GoogleDrive::Session.from_config("config.json")
	directory 	 = drive.folder_by_id("1i9QZ93LHuFnqOft_-GBmjHuIlGRGPNGl")
	current_date = Date.today.strftime("%Y%m%d")
	folder    	 = directory.create_subcollection(current_date)

	task :export_sheet => :environment do
  	# make new folder for remote files
  	directory_name = "remote_files"
		Dir.mkdir(directory_name) unless File.exists?(directory_name)
		# get all sheet links
		sheet_links = Bank::SHEET_LINKS
		# traverse each link one by one and download file inside folder
		sheet_links.each do |link|
			open(link) do |remote_file|
				file_name = remote_file.meta["content-disposition"].gsub("attachment;filename=", "")
				unless File.exists?(Rails.root.join('remote_files', file_name))
					File.open("remote_files/" + file_name, "wb") do |local_file|
						# save file
						local_file.write(remote_file.read)

						# upload data on db through sheet
						# Rake::Task["sheet_operation:import_sheet_data[1]"]
						# rake sheet_operation:import_sheet_data[1]
					end
				end
			end
		end

		# Rake::Task["sheet_operation:upload_on_drive"].invoke
  end

  task :upload_on_drive => :environment do
  	# get all files of specified folder
  	folder_files = Dir.entries("remote_files")
  	folder_files.each do  |file|
  		if file.split(".").last.present?
  			# folder.upload_from_file("remote_files/#{file}", file, convert: false)
  			Rake::Task["sheet_operation:import_sheet_data"].invoke(file)
  		end
  	end
  end

  task :import_sheet_data, [:filename] => :environment do |task, args|
  	file   = File.join(Rails.root.join('remote_files', args[:filename]))
  	xlsx 	 = Roo::Spreadsheet.open(file)
  	sheets = xlsx.sheets
  end
end
