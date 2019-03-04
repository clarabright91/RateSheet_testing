class ErrorLogsController < ApplicationController

  def display_logs
    @error_logs = ErrorLog.where("sheet_name Like ? AND created_at > ?", params[:name], 24.hours.ago)
    render :template => 'error_logs/sheet_logs.html.erb'
  end
end
