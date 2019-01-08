class DashboardController < ApplicationController
  before_action :set_default, :find_base_rate

  def index
    @banks = Bank.all
    if request.post?
      set_variable
      find_base_rate
    end
  end

  def set_default
    @rate_type = "Fixed"
    @interest = "4.375"
    @term = 30
    @day = "30"
    @base_rate = 0.0
  end

  def set_variable
    @rate_type = params[:rate_type]
    @interest = params[:interest].to_s
    @term = params[:term].to_i
    @day = params[:day]
  end

  def find_base_rate
    programs = Program.where(term: @term, rate_type: @rate_type)
    program = programs.find_by_program_name("Fannie Mae 30yr Fixed")
    if program.present?
      @adjustment =  program.adjustments.first
      @adjustment_data = JSON.parse @adjustment.data
      if program.base_rate[@interest].present?
        @base_rate = program.base_rate[@interest][@day]
      end
    end
    return @base_rate
  end

  render :index
end
