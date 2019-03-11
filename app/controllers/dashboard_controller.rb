class DashboardController < ApplicationController
  before_action :set_default

  def index
    @banks = Bank.all
    if params["commit"].present?
      set_variable
      find_base_rate
    end
  end

  def fetch_programs_by_bank
    program_list = []
    if params[:bank_name].present?
      if (params[:bank_name] == "All")
        program_list = Program.all
      else
        program_list = Program.where(bank_name: params[:bank_name])
      end
    end
    render json: {program_list: program_list.map{ |n| {program_name: n.program_name} }}
  end

  def set_default
    @base_rate = 0.0
    @filter_data = {}
    @interest = "4.375"
    @lock_period ="30"
    @loan_size = "High-Balance"
    @loan_type = "Fixed"
    @term = 30
    @ltv = []
    @credit_score = []
    @cltv = []
  end

  def set_variable
    if params[:ltv].present?
      if params[:ltv].include?("-")
        ltv_range = (params[:ltv].split("-").first.to_f..params[:ltv].split("-").last.to_f)
        ltv_range.step(0.01) { |f| @ltv << f }
        @ltv = @ltv.uniq
      elsif params[:ltv].include?("+")
        ltv_range = (params[:ltv].to_f..(params[:ltv].to_f+60))
        ltv_range.step(0.01) { |f| @ltv << f }
        @ltv = @ltv.uniq
      end
    end

    if params[:cltv].present?
      if params[:cltv].include?("-")
        cltv_range = (params[:cltv].split("-").first.to_f..params[:cltv].split("-").last.to_f)
        cltv_range.step(0.01) { |f| @cltv << f }
        @cltv = @cltv.uniq
      elsif params[:cltv].include?("+")
        cltv_range = (params[:cltv].to_f..(params[:cltv].to_f+60))
        cltv_range.step(0.01) { |f| @cltv << f }
        @cltv = @cltv.uniq
      end
    end

    if params[:credit_score].present?
      if  params[:credit_score].include?("-")
        credit_score_range = (params[:credit_score].split("-").first.to_f..params[:credit_score].split("-").last.to_f)
        credit_score_range.step(0.01) { |f| @credit_score << f }
        @credit_score = @credit_score.uniq
      elsif params[:credit_score].include?("+")
          credit_score_range = (params[:credit_score].to_f..(params[:credit_score].to_f+100))
          credit_score_range.step(0.01) { |f| @credit_score << f }
          @credit_score = @credit_score.uniq
      end
    end

    @state = params[:state] if params[:state].present?
    @property_type = params[:property_type] if params[:property_type].present?
    @financing_type = params[:financing_type] if params[:financing_type].present?
    @refinance_option = params[:refinance_option] if params[:refinance_option].present?
    @misc_adjuster = params[:misc_adjuster] if params[:misc_adjuster].present?
    @premium_type = params[:premium_type] if params[:premium_type].present?
    @interest = params[:interest] if params[:interest].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
    @fannie_mae_product = params[:fannie_mae_product] if params[:fannie_mae_product].present?
    @fraddie_mac_product = params[:fraddie_mac_product] if params[:fraddie_mac_product].present?
    @loan_amount = params[:loan_amount].to_i if params[:loan_amount].present?
    @program_category = params[:program_category] if params[:program_category].present?
    @payment_type =  params[:payment_type] if params[:payment_type].present?

    if params[:bank_name].present?
      unless (params[:bank_name] == "All")
        @filter_data[:bank_name] = params[:bank_name]
      end
    end

    if params[:program_name].present?
      unless (params[:program_name] == "All")
        @filter_data[:program_name] = params[:program_name]
      end
    end

    if params[:pro_category].present?
      unless (params[:pro_category] == "All")
        @filter_data[:program_category] = params[:pro_category]
      end
    end

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
        unless params[:arm_advanced] == "All"
          @arm_advanced = params[:arm_advanced]
          @filter_data[:arm_advanced] = params[:arm_advanced]
        end
      end
    end

    if (params[:term].present? && params[:loan_type] != "ARM")
      @term = params[:term].to_i
      @program_term = params[:term].to_i
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
        @filter_data[:high_balance] = true
      end
    end

    if params[:loan_purpose].present?
      @loan_purpose = params[:loan_purpose]
      @filter_data[:loan_purpose] = params[:loan_purpose]
    end
  end

  def find_base_rate
    @program_list = Program.where(@filter_data)
    @program_list2 = []
    if @program_list.present?
      if @program_term.present?
        @program_list.each do |program|
          if (program.term.to_s.length == 2 || program.term.to_s.length == 1)
            if (program.term == @program_term)
              @program_list2 << program
            end
          elsif (program.term.to_s.length == 4)
            first = program.term/100
            last = program.term%100
            if first < last
              if ((first..last).to_a).include?(@program_term)
                @program_list2 << program
              end
            else
              if ((last..first).to_a).include?(@program_term)
                @program_list2 << program
              end
            end
          end
        end
      else
        @program_list2 = @program_list
      end

      @programs =[]
      if @program_list2.present?
        @program_list2.each do |program|
          if(program.base_rate.keys.include?(@interest.to_f.to_s))
            if(program.base_rate[@interest.to_f.to_s].keys.include?(@lock_period))
                @programs << program
            end
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
      :bank_name=> "",
      :adj_points => []
    }
    programs.each do |pro|
      hash_obj[:program_name] = pro.program_name.present? ? pro.program_name : ""
      hash_obj[:sheet_name] = pro.sheet_name.present? ? pro.sheet_name : ""
      hash_obj[:bank_name] = pro.bank_name.present? ? pro.bank_name : ""
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
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key].present?
                      loan_amount_key2 = ''
                      adj.data[first_key].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LTV"
                  begin
                    if adj.data[first_key].present?
                      ltv_key2 = ''
                      adj.data[first_key].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key].present?
                      fico_key2 = ''
                        adj.data[first_key].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "CLTV"
                  begin
                    if adj.data[first_key].present?
                      ltv_key2 = ''
                      adj.data[first_key].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key].keys.first
                      if adj.data[first_key][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

              end
              if key_index==1
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]].present?
                      loan_amount_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]].present?
                      ltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]].present?
                      fico_key2 = ''
                        adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "CLTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]].present?
                      cltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key][adj_key_hash[key_index-1]].keys.first
                      if adj.data[first_key][adj_key_hash[key_index-1]][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][adj_key_hash[key_index-1]][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
              end

              if key_index==2
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      loan_amount_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      ltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      fico_key2 = ''
                        adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "CLTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      cltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                      if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
              end
              if key_index==3
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      loan_amount_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      ltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      fico_key2 = ''
                        adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "CLTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      cltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                      if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
              end
              if key_index==4
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      loan_amount_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      ltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      fico_key2 = ''
                        adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "CLTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      cltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                      if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
              end
              if key_index==5
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      loan_amount_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      ltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      fico_key2 = ''
                        adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "CLTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      cltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                      if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
              end
              if key_index==6
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @fraddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LoanAmount"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      loan_amount_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |loan_amount_key|
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr('$', '').strip
                        end
                        if loan_amount_key.include?("$")
                          loan_amount_key = loan_amount_key.tr(',', '').strip
                        end
                        if (loan_amount_key.include?("Inf") || loan_amount_key.include?("Infinity"))
                          if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i)
                            loan_amount_key2 = loan_amount_key
                            adj_key_hash[key_index] = loan_amount_key
                          end
                        else
                          if loan_amount_key.include?("-")
                            if (loan_amount_key.split("-").first.strip.to_i <= @loan_amount.to_i && @loan_amount.to_i <= loan_amount_key.split("-").second.strip.to_i)
                              loan_amount_key2 = loan_amount_key
                              adj_key_hash[key_index] = loan_amount_key
                            end
                          end
                        end
                      end
                      unless loan_amount_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "PropertyType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@property_type].present?
                      adj_key_hash[key_index] = @property_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "FinancingType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@financing_type].present?
                      adj_key_hash[key_index] = @financing_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                if key_name == "PremiumType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@premium_type].present?
                      adj_key_hash[key_index] = @premium_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
                
                if key_name == "LTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      ltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |ltv_key|
                        if (ltv_key.include?("Any") || ltv_key.include?("All"))
                          ltv_key2 = ltv_key
                          adj_key_hash[key_index] = ltv_key
                        end
                        if ltv_key.include?("-")
                          ltv_key_range =[]
                          if ltv_key.include?("Inf") || ltv_key.include?("Infinity")
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          else
                            (ltv_key.split("-").first.strip.to_f..ltv_key.split("-").last.strip.to_f).step(0.01) { |f| ltv_key_range << f }
                            ltv_key_range = ltv_key_range.uniq
                            if (ltv_key_range & @ltv).present?
                              ltv_key2 = ltv_key
                              adj_key_hash[key_index] = ltv_key
                            end
                          end
                        end
                      end
                      unless ltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "FICO"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      fico_key2 = ''
                        adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |fico_key|
                          if (fico_key.include?("Any") || fico_key.include?("All"))
                            fico_key2 = fico_key
                            adj_key_hash[key_index] = fico_key
                          end
                          if fico_key.include?("-")
                            fico_key_range =[]
                            if fico_key.include?("Inf") || fico_key.include?("Infinity")
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").first.strip.to_f+60).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            else
                              (fico_key.split("-").first.strip.to_f..fico_key.split("-").last.strip.to_f).step(0.01) { |f| fico_key_range << f }
                              fico_key_range = fico_key_range.uniq
                              if (fico_key_range & @credit_score).present?
                                fico_key2 = fico_key
                                adj_key_hash[key_index] = fico_key
                              end
                            end
                          end
                        end
                        unless fico_key2.present?
                          break
                        end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "RefinanceOption"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "MiscAdjuster"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@misc_adjuster].present?
                      adj_key_hash[key_index] = @misc_adjuster
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanSize"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_size].present?
                      adj_key_hash[key_index] = @loan_size
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "CLTV"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].present?
                      cltv_key2 = ''
                      adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.each do |cltv_key|
                        if (cltv_key.include?("Any") || cltv_key.include?("All"))
                          cltv_key2 = cltv_key
                          adj_key_hash[key_index] = cltv_key
                        end
                        if cltv_key.include?("-")
                          cltv_key_range =[]
                          if cltv_key.include?("Inf") || cltv_key.include?("Infinity")
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").first.strip.to_f+60).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          else
                            (cltv_key.split("-").first.strip.to_f..cltv_key.split("-").last.strip.to_f).step(0.01) { |f| cltv_key_range << f }
                            cltv_key_range = cltv_key_range.uniq
                            if (cltv_key_range & @ltv).present?
                              cltv_key2 = cltv_key
                              adj_key_hash[key_index] = cltv_key
                            end
                          end
                        end
                      end
                      unless cltv_key2.present?
                        break
                      end
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state = "All"
                      first_state_key = adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]].keys.first
                      if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][first_state_key].present?
                        adj_key_hash[key_index] = first_state_key
                      else
                        break
                      end
                    else
                      if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@state].present?
                        adj_key_hash[key_index] = @state
                      else
                        break
                      end
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LoanType"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_type].present?
                      adj_key_hash[key_index] = @loan_type
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "Term"
                  begin
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
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end

                if key_name == "LPMI"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"].present?
                      adj_key_hash[key_index] = "true"
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Sheet Name #{adj.sheet_name}"
                  end
                end
              end
            else
              if key_index==0
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==1
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==2
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==3
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==4
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==5
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==6
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI")
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
            end
          end
          adj_key_hash.keys.each do |hash_key, index|
            if hash_key==0 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==1 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==2 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==3 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==4 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==5 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-5]][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
            if hash_key==6 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-6]][adj_key_hash[hash_key-5]][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "-"))
                hash_obj[:adj_points] << point
              end
            end
          end
        end
      end
      @result << hash_obj
        hash_obj = {
        :program_name => "",
        :base_rate => 0.0,
        :sheet_name=> "",
        :bank_name=> "",
        :adj_points => []
      }
    end
  end

  render :index
end