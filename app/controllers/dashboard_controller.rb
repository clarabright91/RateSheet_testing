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
    @credit_score = 740
    @ltv = 81.0
    @cltv = 81.0
    @fico = ""
    @property_type = "Manufactured Home"
    @financing_type = "Subordinate Financing"
    @refinance_option = "Cash Out"
    @misc_adjuster = "CA Escrow Waiver (Full or Taxes Only)"
    @premium_type = "Manufactured Home"
    @state = "CA"
    @loan_size = "High-Balance"
    @loan_type = "Fixed"
    @term = 30
  end

  def set_variable
    @cltv = params[:cltv].to_f if params[:cltv].present?
    @property_type = params[:property_type] if params[:property_type].present?
    @financing_type = params[:financing_type] if params[:financing_type].present?
    @refinance_option = params[:refinance_option] if params[:refinance_option].present?
    @misc_adjuster = params[:misc_adjuster] if params[:misc_adjuster].present?

    @interest = params[:interest] if params[:interest].present?
    @credit_score = params[:credit_score].to_i if params[:credit_score].present?
    @ltv = params[:ltv].to_f if params[:ltv].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
    if params[:loan_type].present?
      @loan_type = params[:loan_type]
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
        @term = params[:term].to_i
      end

      if params[:loan_type] =="Floating" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
        @term = params[:term].to_i
      end

      if  params[:loan_type] =="Variable" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
        @term = params[:term].to_i
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
      @loan_size = params[:loan_size]
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
      @result= []
      if @programs.present?
        find_points_of_the_loan @programs
      end
    end
  end

  def find_points_of_the_loan programs
    hash_obj = {
      :program_name => "",
      :adj_points => []
    }
    programs.each do |pro|      
      hash_obj[:program_name] = pro.program_name
      if pro.adjustments.present?
        pro.adjustments.each do |adj|
          first_key = adj.data.keys.first          
          key_list = first_key.split("/")
          adj_key_hash = {}
          key_list.each_with_index do |key_name, key_index|
            if(Adjustment::INPUT_VALUES.include?(key_name))              
              if key_index==0
                if key_name == "PropertyType"
                  adj.data[first_key][@property_type]
                  adj_key_hash[key_index] = @property_type
                end
                if key_name == "FinancingType"
                  adj.data[first_key][@financing_type]
                  adj_key_hash[key_index] = @financing_type
                end
                if key_name == "PremiumType"
                  adj.data[first_key][@premium_type]
                  adj_key_hash[key_index] = @premium_type
                end
                if key_name == "LTV"                  
                  adj.data[first_key].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][ltv_key]
                        adj_key_hash[key_index] = ltv_key
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end
                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key]["Cash-Out"].present?
                      adj.data[first_key]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key]["Cash Out"].present?
                      adj.data[first_key]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end
                if key_name == "MiscAdjuster"
                  adj.data[first_key][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key].keys.each do |cltv_key|
                    if cltv_key.include?("-")
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end

                if key_name == "Term"
                  term_key = adj.data[first_key].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end
                if key_name == "LPMI"
                  adj.data[first_key]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==1
                if key_name == "PropertyType"
                  adj.data[first_key][adj_key_hash[key_index-1]][@property_type]
                  adj_key_hash[key_index] = @property_type                  
                end
                if key_name == "FinancingType"
                  adj.data[first_key][adj_key_hash[key_index-1]][@financing_type]
                  adj_key_hash[key_index] = @financing_type                  
                end
                if key_name == "PremiumType"
                  adj.data[first_key][adj_key_hash[key_index-1]][@premium_type]
                  adj_key_hash[key_index] = @premium_type                  
                end
                if key_name == "LTV"                  
                  adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-1]][ltv_key]
                        adj_key_hash[key_index] = ltv_key                        
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][adj_key_hash[key_index-1]][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end

                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key][adj_key_hash[key_index-1]]["Cash-Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-1]]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key][adj_key_hash[key_index-1]]["Cash Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-1]]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][adj_key_hash[key_index-1]][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end

                if key_name == "MiscAdjuster"
                  adj.data[first_key][adj_key_hash[key_index-1]][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][adj_key_hash[key_index-1]][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                    if cltv_key.include?("-")                      
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")                      
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][adj_key_hash[key_index-1]][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][adj_key_hash[key_index-1]][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end

                if key_name == "Term"
                  term_key = adj.data[first_key][adj_key_hash[key_index-1]].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end

                if key_name == "LPMI"                  
                  adj.data[first_key][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==2
                if key_name == "PropertyType"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type]
                  adj_key_hash[key_index] = @property_type                  
                end
                if key_name == "FinancingType"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type]
                  adj_key_hash[key_index] = @financing_type                  
                end
                if key_name == "PremiumType"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type]
                  adj_key_hash[key_index] = @premium_type                  
                end
                if key_name == "LTV"                  
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                        adj_key_hash[key_index] = ltv_key                        
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end

                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end

                if key_name == "MiscAdjuster"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                    if cltv_key.include?("-")
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end
                if key_name == "Term"
                  term_key = adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end

                if key_name == "LPMI"                  
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==3
                if key_name == "PropertyType"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type]
                  adj_key_hash[key_index] = @property_type                  
                end
                if key_name == "FinancingType"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type]
                  adj_key_hash[key_index] = @financing_type                  
                end
                if key_name == "PremiumType"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type]
                  adj_key_hash[key_index] = @premium_type                  
                end
                if key_name == "LTV"                                    
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                        adj_key_hash[key_index] = ltv_key                        
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end

                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end

                if key_name == "MiscAdjuster"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                    if cltv_key.include?("-")
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end
                if key_name == "Term"
                  term_key = adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end

                if key_name == "LPMI"
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==4
                if key_name == "PropertyType"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type]
                  adj_key_hash[key_index] = @property_type                  
                end
                if key_name == "FinancingType"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type]
                  adj_key_hash[key_index] = @financing_type                  
                end
                if key_name == "PremiumType"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type]
                  adj_key_hash[key_index] = @premium_type                  
                end
                if key_name == "LTV"                  
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                        adj_key_hash[key_index] = ltv_key                        
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end
                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end
                if key_name == "MiscAdjuster"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                    if cltv_key.include?("-")
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end

                if key_name == "Term"
                  term_key = adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end

                if key_name == "LPMI"
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==5
                if key_name == "PropertyType"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type]
                  adj_key_hash[key_index] = @property_type                  
                end
                if key_name == "FinancingType"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type]
                  adj_key_hash[key_index] = @financing_type                  
                end
                if key_name == "PremiumType"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type]
                  adj_key_hash[key_index] = @premium_type                  
                end
                if key_name == "LTV"                  
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                        adj_key_hash[key_index] = ltv_key                        
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end

                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end

                if key_name == "MiscAdjuster"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                    if cltv_key.include?("-")
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end
                if key_name == "Term"
                  term_key = adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end

                if key_name == "LPMI"
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==6
                if key_name == "PropertyType"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type]
                  adj_key_hash[key_index] = @property_type                  
                end
                if key_name == "FinancingType"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type]
                  adj_key_hash[key_index] = @financing_type                  
                end
                if key_name == "PremiumType"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type]
                  adj_key_hash[key_index] = @premium_type                  
                end
                if key_name == "LTV"                  
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                    if ltv_key.include?("-")
                      if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                        adj_key_hash[key_index] = ltv_key                        
                      end
                    end
                  end
                end
                if key_name == "FICO"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                    if fico_key.include?("-")
                      if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                        adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                        adj_key_hash[key_index] = fico_key                        
                      end
                    end
                  end
                end
                if key_name == "RefinanceOption"                  
                  if (@refinance_option == "Cash Out" || @refinance_option == "Cash-Out")
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash-Out"]
                      adj_key_hash[key_index] = "Cash-Out"
                    end
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"].present?
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["Cash Out"]
                      adj_key_hash[key_index] = "Cash Out"
                    end
                  else
                    adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@refinance_option]
                    adj_key_hash[key_index] = @refinance_option                    
                  end
                end
                if key_name == "MiscAdjuster"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster]
                  adj_key_hash[key_index] = @misc_adjuster                  
                end
                if key_name == "LoanSize"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size]
                  adj_key_hash[key_index] = @loan_size                  
                end
                if key_name == "CLTV"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                    if cltv_key.include?("-")
                      if (cltv_key.split("-").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("-").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                    if cltv_key.include?("<=")
                      if (cltv_key.split("<=").first.strip.to_f < @cltv.to_f && @cltv.to_f <= cltv_key.split("<=").second.strip.to_f)
                        adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][cltv_key]
                        adj_key_hash[key_index] = cltv_key                        
                      end
                    end
                  end
                end
                if key_name == "State"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state]
                  adj_key_hash[key_index] = @state                  
                end
                if key_name == "LoanType"                  
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type]
                  adj_key_hash[key_index] = @loan_type                  
                end

                if key_name == "Term"
                  term_key = adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                  if term_key.include?("Inf") || term_key.include?("Infinite")                    
                    if (term_key.split("-").first.strip.to_i <= @term)
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  elsif term_key.include?("-")                    
                    if (term_key.split("-").first.strip.to_i <= @term && @term <= term_key.split("-").second.strip.to_i)
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][term_key]
                      adj_key_hash[key_index] = term_key                      
                    else
                      break
                    end
                  end
                end

                if key_name == "LPMI"
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
            else
              if key_index==0
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==1
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==2
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==3
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==4
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==5
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
              if key_index==6
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc")                  
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"                  
                end
              end
            end
          end
          adj_key_hash.keys.each do |hash_key, index|            
          end
        end
      end
    end
  end

  render :index
end
