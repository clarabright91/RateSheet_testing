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
    @lock_period ="30"
    @credit_score = 740
    @ltv = 81.0
    @cltv = 81.0
    @loan_size = "High-Balance"
    @loan_type = "Fixed"
    @term = 30
  end

  def set_variable
    @cltv = params[:cltv].to_f if params[:cltv].present?
    @state = params[:state].to_f if params[:state].present?
    @property_type = params[:property_type] if params[:property_type].present?
    @financing_type = params[:financing_type] if params[:financing_type].present?
    @refinance_option = params[:refinance_option] if params[:refinance_option].present?
    @misc_adjuster = params[:misc_adjuster] if params[:misc_adjuster].present?
    @premium_type = params[:premium_type] if params[:premium_type].present?

    @interest = params[:interest] if params[:interest].present?
    @credit_score = params[:credit_score].to_i if params[:credit_score].present?
    @ltv = params[:ltv].to_f if params[:ltv].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
    @fannie_mae_product = params[:fannie_mae_product] if params[:fannie_mae_product].present?
    @fraddie_mac_product = params[:fraddie_mac_product] if params[:fraddie_mac_product].present?
    @loan_amount = params[:loan_amount] if params[:loan_amount].present?
    @program_category = params[:program_category] if params[:program_category].present?

    if params[:loan_type].present?
      @loan_type = params[:loan_type]
      @filter_data[:loan_type] = params[:loan_type]
      if params[:loan_type] =="ARM" && params[:arm_basic].present?
        @arm_basic = params[:arm_basic]
        if params[:arm_basic].include?("/")
          @filter_data[:arm_basic] = params[:arm_basic].split("/").first
        end
        if params[:arm_basic].include?("-")
          @filter_data[:arm_basic] = params[:arm_basic].split("-").first
        end
      end
      if params[:loan_type] =="ARM" && params[:arm_advanced].present?
        @arm_advanced = params[:arm_advanced]
        @filter_data[:arm_advanced] = params[:arm_advanced]
      end
    end

    if (params[:term].present? && params[:loan_type] != "ARM")
      @filter_data[:term] = params[:term].to_i
      @term = params[:term].to_i
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
      @loan_purpose = params[:loan_purpose]
      @filter_data[:loan_purpose] = params[:loan_purpose]
    end
  end

  def find_base_rate
    @program_list = Program.where(@filter_data) 
    if @program_list.present?
      @programs =[]
      @program_list.each do |program|
        if(program.base_rate.keys.include?(@interest.to_f.to_s))
          if(program.base_rate[@interest.to_f.to_s].keys.include?(@lock_period))
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
      :base_rate => 0.0,
      :sheet_name=> "",
      :adj_points => []
    }
    programs.each do |pro|
      hash_obj[:program_name] = pro.program_name.present? ? pro.program_name : ""
      hash_obj[:sheet_name] = pro.sheet_name.present? ? pro.sheet_name : ""
      hash_obj[:base_rate] = pro.base_rate[@interest.to_f.to_s][@lock_period].present? ? pro.base_rate[@interest.to_f.to_s][@lock_period] : 0.0
      if pro.adjustments.present?
        pro.adjustments.each do |adj|
          first_key = adj.data.keys.first          
          key_list = first_key.split("/")
          adj_key_hash = {}
          key_list.each_with_index do |key_name, key_index|
            if(Adjustment::INPUT_VALUES.include?(key_name))
              if key_index==0
                if key_name == "LockDay"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @lock_period
                  else
                    break
                  end
                end
                if key_name == "ArmBasic"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "ArmAdvanced"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "FannieMaeProduct"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "FreddieMacProduct"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "LoanPurpose"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "LoanAmount"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "ProgramCategory"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end

                if key_name == "PropertyType"
                  if adj.data[first_key][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                end
                if key_name == "LTV"
                  if adj.data[first_key].present?
                    adj.data[first_key].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key].present?
                      adj.data[first_key].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                  if adj.data[first_key]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
              if key_index==1
                if key_name == "PropertyType"
                  if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                                  
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][adj_key_hash[key_index-1]][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                                  
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][adj_key_hash[key_index-1]][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                                  
                end
                if key_name == "LTV"
                  if adj.data[first_key][adj_key_hash[key_index-1]].present?
                    adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][adj_key_hash[key_index-1]][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                                                
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key][adj_key_hash[key_index-1]].present?
                      adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][adj_key_hash[key_index-1]][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][adj_key_hash[key_index-1]][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][adj_key_hash[key_index-1]][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key][adj_key_hash[key_index-1]].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][adj_key_hash[key_index-1]][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][adj_key_hash[key_index-1]][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                                
                  if adj.data[first_key][adj_key_hash[key_index-1]]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
              if key_index==2
                if key_name == "PropertyType"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                                  
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                                  
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                                  
                end
                if key_name == "LTV"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                    adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                                                
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                                
                  if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
              if key_index==3
                if key_name == "PropertyType"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                                  
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                                  
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                                  
                end
                if key_name == "LTV"
                if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                    adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                                                
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                  if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
              if key_index==4
                if key_name == "PropertyType"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                                  
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                                  
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                                  
                end
                if key_name == "LTV"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                    adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                                                
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                  if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
              if key_index==5
                if key_name == "PropertyType"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                                  
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                                  
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                                  
                end
                if key_name == "LTV"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                    adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                                                
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                  if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
              if key_index==6
                if key_name == "PropertyType"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                    adj_key_hash[key_index] = @property_type
                  else
                    break
                  end
                                  
                end
                if key_name == "FinancingType"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                    adj_key_hash[key_index] = @financing_type
                  else
                    break
                  end
                                  
                end
                if key_name == "PremiumType"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                    adj_key_hash[key_index] = @premium_type
                  else
                    break
                  end
                                  
                end
                if key_name == "LTV"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                    adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                      if ltv_key.include?("-")
                        if (ltv_key.split("-").first.strip.to_f < @ltv.to_f && @ltv.to_f <= ltv_key.split("-").second.strip.to_f)
                          adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][ltv_key]
                          adj_key_hash[key_index] = ltv_key
                                                
                        end
                      end
                    end
                  else
                    break
                  end
                end
                if key_name == "FICO"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                        if fico_key.include?("-")
                          if (fico_key.split("-").first.strip.to_i <= @credit_score && fico_key.split("-").second.strip.to_i >= @credit_score)
                            adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][fico_key]
                            adj_key_hash[key_index] = fico_key                                                  
                          end
                        end
                      end
                  else
                    break
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
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                    adj_key_hash[key_index] = @misc_adjuster
                  else
                    break
                  end
                                  
                end
                if key_name == "LoanSize"                  
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                    adj_key_hash[key_index] = @loan_size
                  else
                    break
                  end
                                  
                end
                if key_name == "CLTV"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
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
                end
                if key_name == "State"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                    adj_key_hash[key_index] = @state
                  else
                    break
                  end
                end
                if key_name == "LoanType"
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                    adj_key_hash[key_index] = @loan_type
                  else
                    break
                  end
                                  
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
                  if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                    adj_key_hash[key_index] = "true"
                  else
                    break
                  end
                end
              end
            else              
              if key_index==0
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==1
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==2
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==3
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==4
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==5
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==6
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo")
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
            end
          end
          adj_key_hash.keys.each do |hash_key, index|            
            if hash_key==0 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==1 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==2 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==3 && adj_key_hash.keys.count-1==hash_key              
              point = adj.data[first_key][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==4 && adj_key_hash.keys.count-1==hash_key              
              point = adj.data[first_key][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==5 && adj_key_hash.keys.count-1==hash_key              
              point = adj.data[first_key][adj_key_hash[hash_key-5]][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==6 && adj_key_hash.keys.count-1==hash_key              
              point = adj.data[first_key][adj_key_hash[hash_key-6]][adj_key_hash[hash_key-5]][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A"))
                hash_obj[:adj_points] << point
              end
            end
          end
        end
        @result << hash_obj
        hash_obj = {
        :program_name => "",
        :base_rate => 0.0,
        :sheet_name=> "",
        :adj_points => []
      }
      end      
    end    
  end

  render :index
end
