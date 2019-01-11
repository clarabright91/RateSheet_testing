class DashboardController < ApplicationController
  before_action :set_default

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
    @rate_type = "None"
    @interest = "4.375"
    @term = 30
    @lock_period = "30"
    @base_rate = 0.0

    @fha = false
    @va = false
    @usda = false
    @conforming = false
    @Jumbo = false
    @high_balance = false

  end

  def set_variable
    @rate_type = params[:rate_type] if params[:rate_type].present?
    @interest = params[:interest] if params[:interest].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
    
    if params[:rate_type] =="ARM" && params[:rate_type].present?
      @term = params[:term_arm].to_i if params[:term].present?
    else
      @term = params[:term].to_i if params[:term].present?
    end
    
    @gov_sheet = params[:gov] if params[:gov].present?
    if @gov_sheet.present?
      if @gov_sheet == "FHA"
        @fha = true
      elsif @gov_sheet == "VA"
        @va = true
      elsif @gov_sheet == "USDA"
        @usda = true
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
     @programs = Program.where(term: @term, rate_type: @rate_type, va:@va, fha: @fha, usda: @usda, conforming: @conforming, jumbo_high_balance: @high_balance)
      if @programs.present?
        program = @programs.first
          if program.base_rate[program.base_rate.keys.first][@interest.to_f.to_s].present?
            @base_rate = program.base_rate[program.base_rate.keys.first][@interest.to_f.to_s][@lock_period]
          else
            flash[:error] = "Not find any interest rate for this situation"
          end
      else
        flash[:error] = "Not find any program for this situation"
      end
    return @base_rate
  end

  # render :index
end
