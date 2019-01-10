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
    @sheet = "Cover Zone 1"
    @gov_sheet = "FHA"
    @loan_limit_type = "High-Balance"
    @fha = false
    @va = false
    @usda = false

    @non_conforming = false
    @conforming = false
    @Jumbo = false
    @high_balance = false
  end

  def set_variable
    @rate_type = params[:rate_type] if params[:rate_type].present?
    @interest = params[:interest] if params[:interest].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
    @sheet = params[:sheet] if params[:sheet].present?
    if params[:rate_type] =="ARM" && params[:rate_type].present?
      @term = params[:term_arm].to_i if params[:term].present?
    else
      @term = params[:term].to_i if params[:term].present?
    end
    
    @gov_sheet = params[:gov] if params[:gov].present?
    if @gov_sheet.present?
      if @gov_sheet == "FHA"
        @fha = true
        @va = false
        @usda = false
      elsif @gov_sheet == "VA"
        @va = true
        @fha = false
        @usda = false
      elsif @gov_sheet == "USDA"
        @usda = true
        @va = false
        @fha = false
      else
        @usda = false
        @va = false
        @fha = false
      end
    end

    @loan_limit_type = params[:loan_limit_type] if params[:loan_limit_type].present?
    if @loan_limit_type.present?
      if @loan_limit_type == "Non-Conforming"
        @non_conforming = true
        @conforming = false
        @Jumbo = false
        @high_balance = false
      elsif @loan_limit_type == "Conforming"
        @non_conforming = false
        @conforming = true
        @Jumbo = false
        @high_balance = false
      elsif @loan_limit_type == "Jumbo"
        @non_conforming = false
        @conforming = false
        @Jumbo = true
        @high_balance = false
      elsif @loan_limit_type == "High-Balance"
        @non_conforming = false
        @conforming = false
        @Jumbo = false
        @high_balance = true
      else
        @non_conforming = false
        @conforming = false
        @Jumbo = false
        @high_balance = false
      end
    end
  end

  def find_base_rate
    
    programs_list = Program.where(sheet_name: @sheet)
     programs = programs_list.where(term: 30, rate_type: "Fixed", va:@va, fha: @fha, usda: @usda)
    # if program.present?
    #   # Adjustment::MAIN_KEYS.key("FinancingType/LTV/CLTV/FICO")
    #   @adjustment =  program.adjustments.first
    #   @adjustment_data = JSON.parse @adjustment.data
    #   @adjustment_data.keys.each do |adjustment_key|
    #     if adjustment_key =="Conforming/RateType/Term/LTV/FICO"
    #         @adjustment_data[adjustment_key]
    #     end
    #     @adjustment_data[adjustment_key]
    #   end

    #   if program.base_rate[@interest].present?
    #     @base_rate = program.base_rate[@interest][@lock_period]
    #   end
    # end
    return @base_rate
  end

  render :index
end
