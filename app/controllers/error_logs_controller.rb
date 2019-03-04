class ErrorLogsController < ApplicationController

	def display_logs
		@error_logs = ErrorLog.all#ErrorLog.where("sheet_name Like ? AND created_at > ?", "Conforming Fixed Rate", 24.hours.ago)
		render :template => 'error_logs/sheet_logs.html.erb'
	end
end
