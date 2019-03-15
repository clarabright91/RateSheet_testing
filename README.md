# README

This README would normally document whatever steps are necessary to get the
application up and running.

Things you may want to cover:

* Ruby version: 2.5.0

* System dependencies
  * Rails version: 5.2.2
* Configuration
  **
* Database creation
  rake db:create
  rake db:migrate
* Database initialization

* How to run the test suite

* Services (job queues, cache servers, search engines, etc.)
  * [Click here for Sidekiq configuration](http://ruby-journal.com/how-to-integrate-sidekiq-with-activejob/)
  * [Click here for Redis configuration](https://www.digitalocean.com/community/tutorials/how-to-install-and-use-redis)
* Deployment instructions

* Configuration and working of google drive script:
  * Configuration
    * Add gem 'google_drive' for more deatils follow (https://github.com/gimite/google-drive-ruby) link.
    * bundle install
    * Setup client id and client secret follow (https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md) links step.
    * Create a config.json file inside initializers
      * Create one hash
      * Add "client_id" as key and add your generated client key as a value
    * Access your drive through GoogleDrive::Session.from_config("config.json")

  * Working
    * Run rake sheet_operation:export_sheet
      * This Command will execute below steps
        * This will create new remote files folder
        * Find all sheet links
        * Traverse each link one by one and download file inside folder
        * Check file present or not in folder
        * Download file through url
        * Extract file and call controller method
        * Than upload file on google drive
* ...
