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
    @lock_period = "30"
    @base_rate = 0.0
    @program_name = "Fannie Mae 30yr Fixed"
  end

  def set_variable
    @rate_type = params[:rate_type]
    @interest = params[:interest].to_s
    @term = params[:term].to_i
    @lock_period = params[:lock_period]
  end

  def find_base_rate
    programs = Program.where(term: @term, rate_type: @rate_type)
    program = programs.find_by_program_name("Fannie Mae 30yr Fixed")
    if program.present?
      # Adjustment::MAIN_KEYS.key("FinancingType/LTV/CLTV/FICO")
      @adjustment =  program.adjustments.first
      @adjustment_data = JSON.parse @adjustment.data
      @adjustment_data.keys.each do |adjustment_key|
        if adjustment_key =="Conforming/RateType/Term/LTV/FICO"
            @adjustment_data[adjustment_key]
        end
        @adjustment_data[adjustment_key]
      end

      if program.base_rate[@interest].present?
        @base_rate = program.base_rate[@interest][@lock_period]
      end
    end
    return @base_rate
  end

  render :index
end
