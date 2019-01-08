class DashboardController < ApplicationController
  before_action :set_default, :find_programs

  def index
    @banks = Bank.all
    if request.post?
      set_variable
      find_programs
    end
  end

  def set_default
    @loan_type = 1
    @interest = "4.375"
    @term = "10"
    @day = "30"
  end

  def set_variable
    @loan_type = params[:loan_type].to_i
    @interest = params[:interest].to_s
    @term = params[:term].to_s
    @day = params[:day].to_s
  end

  def find_programs
    programs =  Program.where(term: @term,loan_type: @loan_type)
    @program = programs.find_by_title("VA 30 Yr Fixed")
    @program.base_rate[@interest][@day]
  end

  render :index
end
