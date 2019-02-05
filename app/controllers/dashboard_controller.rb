class DashboardController < ApplicationController
  before_action :set_default

  def index
    @banks = Bank.all
    if params["commit"].present?
      set_variable
      find_base_rate
    end
  end

  def set_default
    @base_rate = 0.0
    @filter_data = {}
    @interest = "4.375"
    @lock_period =30
    @credit_score = "740-759"
  end

  def set_variable
    @interest = params[:interest] if params[:interest].present?
    @credit_score = params[:credit_score] if params[:credit_score].present?

    @lock_period = params[:lock_period] if params[:lock_period].present?

    if params[:loan_type].present?
      @filter_data[:loan_type] = params[:loan_type]
      if params[:loan_type] =="ARM" && params[:arm_basic].present?
        if params[:arm_basic] =="11"
          @filter_data[:arm_basic] = 10
          @filter_data[:arm_advanced] = "3-2-5"
        elsif params[:arm_basic] =="12"
          @filter_data[:arm_basic] = 5
          @filter_data[:arm_advanced] = "3-2-5"
        else
          @filter_data[:arm_basic] = params[:arm_basic].to_i
        end
      end

      if params[:loan_type] =="Fixed" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
      end

      if params[:loan_type] =="Floating" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
      end

      if  params[:loan_type] =="Variable" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
      end
    end

    if params[:fannie_options].present?
      if params[:fannie_options] == "Fannie Mae"
        @filter_data[:fannie_mae] = true
      elsif params[:fannie_options] == "Fannie Mae Home Ready"
        @filter_data[:fannie_mae_home_ready] = true
      elsif params[:fannie_options] == "Fannie Mac"
        @filter_data[:freddie_mac] = true
      elsif params[:fannie_options] == "Fannie Mae freddie_mac_home_possible Possible"
        @filter_data[:freddie_mac_home_possible] = true
      end
    end

    if params[:gov].present?
      if params[:gov] == "FHA"
        @filter_data[:fha] = true
      elsif params[:gov] == "VA"
        @filter_data[:va] = true
      elsif params[:gov] == "USDA"
        @filter_data[:usda] = true
      elsif params[:gov] == "FHA" || params[:gov] == "VA" || params[:gov] == "USDA"
        @filter_data[:streamline] = true
        @filter_data[:full_doc] = true
      end
    end

    if params[:loan_size].present?
      if params[:loan_size] == "Non-Conforming"
        @filter_data[:conforming] = false
      elsif params[:loan_size] == "Conforming"
        @filter_data[:conforming] = true
      elsif params[:loan_size] == "Jumbo"
        @filter_data[:Jumbo] = true
      elsif params[:loan_size] == "High-Balance"
        @filter_data[:jumbo_high_balance] = true
      end
    end

    if params[:loan_purpose].present?
      @filter_data[:loan_purpose] = params[:loan_purpose]
    end
  end

  def find_base_rate
    @program_list = Program.where(@filter_data)
    if @program_list.present?
      @programs =[]
      @program_list.each do |program|
        if(program.base_rate[program.base_rate.keys.first].keys.include?(@interest.to_f.to_s))
          if(program.base_rate[program.base_rate.keys.first][@interest.to_f.to_s].keys.include?(@lock_period))
              @programs << program
          end
        end
      end
    end
  end

  render :index
end
