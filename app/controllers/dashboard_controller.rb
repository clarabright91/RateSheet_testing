class DashboardController < ApplicationController
  before_action :set_default, :find_base_rate

  def index
    @banks = Bank.all
    if request.post?
      set_variable
      find_base_rate
    end
  end

  def fetch_program_list
    
  end

  def set_default
    @rate_type = "Fixed"
    @interest = "4.375"
    @term = 30
    @lock_period = "30"
    @base_rate = 0.0
    # @program_name = "Fannie Mae 30yr Fixed"
    # @sheet = "Cover Zone 1"

    @fha = false
    @va = false
    @usda = false
    @conforming = false
    @Jumbo = false
    @high_balance = false

    @fannie_mae = false
    @fannie_mae_home_ready = false
    @freddie_mac = false
    @freddie_mac_home_possible = false

  end

  def set_variable
    @rate_type = params[:rate_type] if params[:rate_type].present?
    @interest = params[:interest] if params[:interest].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
    # @sheet = params[:sheet] if params[:sheet].present?
    
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
        @conforming = false
      elsif @loan_limit_type == "Conforming"
        @conforming = true
      elsif @loan_limit_type == "Jumbo"
        @Jumbo = true
      elsif @loan_limit_type == "High-Balance"
        @high_balance = true
      elsif @loan_limit_type == "Fannie Mae"
        @fannie_mae = true
      elsif @loan_limit_type == "Fannie Mae Home Ready"
        @fannie_mae_home_ready = true
      elsif @loan_limit_type == "Freddie Mac"
        @freddie_mac = true
      elsif @loan_limit_type == "Freddie Mac Home Possible"
        @freddie_mac_home_possible = true
      else
        @conforming = false
        @Jumbo = false
        @high_balance = false
        @fannie_mae = false
        @fannie_mae_home_ready = false
        @freddie_mac = false
        @freddie_mac_home_possible = false
      end
    end
  end

  def find_base_rate
    # binding.pry
    
     programs = Program.where(term: @term, rate_type: @rate_type, va:@va, fha: @fha, usda: @usda, jumbo_high_balance: @high_balance, conforming: @conforming, fannie_mae: @fannie_mae, fannie_mae_home_ready: @fannie_mae_home_ready,freddie_mac: @freddie_mac, freddie_mac_home_possible: @freddie_mac_home_possible)

     if programs.present?
       program = programs.first
        @adjustment =  program.adjustments.first
        @adjustment_data = JSON.parse @adjustment.data
        if program.base_rate[@interest].present?
          @base_rate = program.base_rate[@interest][@lock_period]
        end
     end

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
