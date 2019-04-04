class DashboardController < ApplicationController
  before_action :set_default

  def index
    @banks = Bank.all
    @all_banks_name = @banks.pluck(:name)
    if params["commit"].present?
      set_variable
      find_base_rate
    end
    fetch_programs_by_bank(true)
  end

  def banks
    @banks = Bank.all
  end

  def fetch_programs_by_bank(html_type=false)
    @all_programs = Program.all
    @program_names = @all_programs.pluck(:program_name).uniq.compact
    @loan_categories = @all_programs.pluck(:loan_category).uniq.compact
    @program_categories = @all_programs.pluck(:program_category).uniq.compact

    if params[:bank_name].present?
      if (params[:bank_name] == "All")
        @program_names = @all_programs.pluck(:program_name).uniq.compact
        @loan_categories = @all_programs.pluck(:loan_category).uniq.compact
        @program_categories = @all_programs.pluck(:program_category).uniq.compact
      else
        @all_programs = @all_programs.where(bank_name: params[:bank_name])
        @program_names = @all_programs.pluck(:program_name).uniq.compact
        @loan_categories = @all_programs.pluck(:loan_category).uniq.compact
        @program_categories = @all_programs.pluck(:program_category).uniq.compact
      end
    end
    if params[:loan_category].present?
      if (params[:loan_category] == "All")
        @program_names = @all_programs.pluck(:program_name).uniq.compact
        @program_categories = @all_programs.pluck(:program_category).uniq.compact
      else
        @all_programs = @all_programs.where(loan_category: params[:loan_category])
        @program_names = @all_programs.pluck(:program_name).uniq.compact
        @program_categories = @all_programs.pluck(:program_category).uniq.compact
      end
    end
    if params[:pro_category].present?
      if (params[:pro_category] == "All" || params[:pro_category] == "No Category")
        @program_names = @all_programs.pluck(:program_name).uniq.compact
      else
        @all_programs = @all_programs.where(program_category: params[:pro_category])
        @program_names = @all_programs.pluck(:program_name).uniq.compact
      end
    end

    if @program_categories.present?
      @program_categories.prepend(["All"])
    else
      @program_categories << "No Category"
    end

    render json: {program_list: @program_names.map{ |lc| {name: lc}}, loan_category_list: @loan_categories.map{ |lc| {name: lc}}, pro_category_list: @program_categories.map{ |lc| {name: lc}}} unless html_type
  end

  def set_default
    @base_rate = 0.0
    @filter_data = {}
    @filter_not_nil = {}
    @interest = "4.375"
    @lock_period ="30"
    @loan_size = "High-Balance"
    @loan_type = "Fixed"
    @term = 30
    @ltv = []
    @credit_score = []
    @cltv = []
    @fannie_mae_product = "HomeReady"
    @freddie_mac_product = "Home Possible"
  end

  def modified_ltv_cltv_credit_score
    %w[ltv cltv credit_score].each do |key|
      array_data = []
      key_value = params[key.to_sym]
      if key_value.present?
        if key_value.include?("-")
          key_range = (key_value.split("-").first.to_f..key_value.split("-").last.to_f)
          key_range.step(0.01) { |f| array_data << f }
          instance_variable_set("@#{key}", array_data.uniq)
        elsif key_value.include?("+")
          score = key.eql?('credit_score') ? 100 : 60
          key_range = (key_value.to_f..(key_value.to_f+score))
          key_range.step(0.01) { |f| array_data << f }
          instance_variable_set("@#{key}", array_data.uniq)
        end
      end
    end
  end

  def modified_condition
    %w[fannie_mae_product freddie_mac_product bank_name program_name pro_category loan_category loan_purpose term].each do |key|
      key_value = params[key.to_sym]
      if key_value.present?
        unless (key_value == "All")
          if (key == "pro_category")
            unless (key_value == "No Category")
              @filter_data[:program_category] = key_value
            end
          else
            @filter_data[key.to_sym] = key_value
          end
          if %w[fannie_mae_product freddie_mac_product loan_size].include?(key)
            instance_variable_set("@#{key}", key_value)
          end
          if %w[term].include?(key)
            if (params[:loan_type] != "ARM")
              instance_variable_set("@#{key}", key_value)
              @program_term = key_value.to_i
            end
          end
        else
          if %w[fannie_mae_product freddie_mac_product loan_purpose loan_size].include?(key)
            @filter_not_nil[key.to_sym] = nil
          end
        end
      end
    end
  end

  def modified_true_condition
    %w[fannie_mae freddie_mac du lp fha va usda streamline full_doc].each do |key|
      key_value = params[key.to_sym]
      if key_value.present?
        @filter_data[key.to_sym] = true
      end
    end
  end

  def modified_variables
    %w[state property_type financing_type refinance_option refinance_option misc_adjuster premium_type interest lock_period loan_amount program_category payment_type].each do |key|
      key_value = params[key.to_sym]
      key_value = key_value.to_i if key_value.present? && key.eql?('loan_amount')
      instance_variable_set("@#{key}", key_value) if key_value.present?
    end
  end

  def set_variable
    modified_ltv_cltv_credit_score
    modified_condition
    modified_true_condition
    modified_variables

    if params[:loan_type].present?
      @loan_type = params[:loan_type]
      if params[:loan_type] == "All"
        @filter_not_nil[:loan_type] = nil
      else
        @filter_data[:loan_type] = params[:loan_type]
        if params[:loan_type] =="ARM" && params[:arm_basic].present?
          @arm_basic = params[:arm_basic]
          unless (params[:arm_basic] == "All")
            if params[:arm_basic].include?("/")
              @filter_data[:arm_basic] = params[:arm_basic].split("/").first
            end
            if params[:arm_basic].include?("-")
              @filter_data[:arm_basic] = params[:arm_basic].split("-").first
            end
          end
        end
        if params[:loan_type] =="ARM" && params[:arm_advanced].present?
          @arm_advanced = params[:arm_advanced]
          unless params[:arm_advanced] == "All"
            @filter_data[:arm_advanced] = params[:arm_advanced]
          end
        end
      end
    end
  end

  def find_base_rate
    @program_list = Program.where(@filter_data)
    @program_list = @program_list.where.not(@filter_not_nil)
    @program_list2 = []
    if @program_list.present?
      if @program_term.present?
        @program_list = @program_list.where.not(term:nil)
        @program_list.each do |program|
           pro_term = program.term
          if (pro_term.to_s.length <=2 )
            if (pro_term == @program_term)
              @program_list2 << program
            end
          else
            first = pro_term/100
            last = pro_term%100
            term_arr = []
            if first < last
              term_arr = (first..last).to_a
            else
              term_arr = (last..first).to_a
            end
            if term_arr.include?(@program_term)
              @program_list2 << program
            end
          end
        end
      else
        @program_list2 = @program_list
      end

      if @program_list2.present?
        @program_list3 = []
        if params[:loan_size].present?
          @loan_size = params[:loan_size]
          @program_list2 = @program_list2.map{ |pro| pro if pro.loan_size!=nil}.compact
          @program_list2.each do |pro|
            if(pro.loan_size.split("and").map{ |l| l.strip }.include?(params[:loan_size]))
              @program_list3 << pro
            end
          end
        else
          @program_list3 = @program_list2
        end
      end

      @programs =[]
      if @program_list3.present?
        @program_list3.each do |program|
          if program.base_rate.present?
            if(program.base_rate.keys.include?(@interest.to_f.to_s))
              if(program.base_rate[@interest.to_f.to_s].keys.include?(@lock_period))
                  @programs << program
              end
            elsif (program.base_rate.keys.include?(@interest.to_s))
              if(program.base_rate[@interest.to_s].keys.include?(@lock_period))
                  @programs << program
              end
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
      :bank_name => "",
      :loan_category => "",
      :program_category => "",
      :program_name => "",
      :base_rate => 0.0,
      :adj_points => [],
      :adj_primary_key => [],
      :final_rate => []
    }
    programs.each do |pro|
      hash_obj[:bank_name] = pro.bank_name.present? ? pro.bank_name : ""
      hash_obj[:loan_category] = pro.loan_category.present? ? pro.loan_category : ""
      hash_obj[:program_category] = pro.program_category.present? ? pro.program_category : ""
      hash_obj[:program_name] = pro.program_name.present? ? pro.program_name : ""

      if (pro.base_rate.present? || pro.base_rate[@interest.to_f.to_s].present? || pro.base_rate[@interest.to_s].present?)
        if (pro.base_rate[@interest.to_f.to_s][@lock_period].present?)
          hash_obj[:base_rate] = pro.base_rate[@interest.to_f.to_s][@lock_period]
        elsif (pro.base_rate[@interest.to_s][@lock_period].present?)
          hash_obj[:base_rate] = pro.base_rate[@interest.to_s][@lock_period]
        else
          hash_obj[:base_rate] = 0.0
        end
      end
      program_adjustments = pro.adjustments
      if program_adjustments.present?
        program_adjustments.each do |adj|
          first_key = adj.data.keys.first
          key_list = first_key.split("/")
          adj_key_hash = {}

          key_list.each_with_index do |key_name, key_index|
            if(Adjustment::INPUT_VALUES.include?(key_name))
              if key_index==0
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

              end
              if key_index==1
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-1]][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
              end

              if key_index==2
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
              end
              if key_index==3
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
              end
              if key_index==4
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
              end
              if key_index==5
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
              end
              if key_index==6
                if key_name == "LockDay"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@lock_period].present?
                      adj_key_hash[key_index] = @lock_period
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmBasic"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_basic].present?
                      adj_key_hash[key_index] = @arm_basic
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "ArmAdvanced"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@arm_advanced].present?
                      adj_key_hash[key_index] = @arm_advanced
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FannieMaeProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@fannie_mae_product].present?
                      adj_key_hash[key_index] = @fannie_mae_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "FreddieMacProduct"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@freddie_mac_product].present?
                      adj_key_hash[key_index] = @freddie_mac_product
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
                if key_name == "LoanPurpose"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@loan_purpose].present?
                      adj_key_hash[key_index] = @loan_purpose
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "ProgramCategory"
                  begin
                    if adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]][@program_category].present?
                      adj_key_hash[key_index] = @program_category
                    else
                      break
                    end
                  rescue Exception
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end

                if key_name == "State"
                  begin
                    if @state == "All"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
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
                    puts "Adjustment Error: Adjustment Id: #{adj.id}, Adjustment Primary Key: #{first_key}, Key Name: #{key_name}, Loan Category: #{adj.loan_category}"
                  end
                end
              end
            else
              if key_index==0
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==1
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==2
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==3
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==4
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==5
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
              if key_index==6
                if (key_name == "HighBalance" || key_name == "Conforming" || key_name == "FannieMae" || key_name == "FannieMaeHomeReady" || key_name == "FreddieMac" || key_name == "FreddieMacHomePossible" || key_name == "FHA" || key_name == "VA" || key_name == "USDA" || key_name == "StreamLine" || key_name == "FullDoc" || key_name == "Jumbo" || key_name == "FHLMC" || key_name == "LPMI" || key_name == "EPMI" || key_name == "FNMA")
                  adj.data[first_key][adj_key_hash[key_index-6]][adj_key_hash[key_index-5]][adj_key_hash[key_index-4]][adj_key_hash[key_index-3]][adj_key_hash[key_index-2]][adj_key_hash[key_index-1]]["true"]
                  adj_key_hash[key_index] = "true"
                end
              end
            end
          end
          adj_key_hash.keys.each do |hash_key, index|
            if hash_key==0 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
            if hash_key==1 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
            if hash_key==2 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
            if hash_key==3 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
            if hash_key==4 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
            if hash_key==5 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-5]][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
            if hash_key==6 && adj_key_hash.keys.count-1==hash_key
              point = adj.data[first_key][adj_key_hash[hash_key-6]][adj_key_hash[hash_key-5]][adj_key_hash[hash_key-4]][adj_key_hash[hash_key-3]][adj_key_hash[hash_key-2]][adj_key_hash[hash_key-1]][adj_key_hash[hash_key]]
              if (((point.is_a? Float) || (point.is_a? Integer) || (point.is_a? String)) && (point != "N/A") && (point != "n/a") && (point != "NA") && (point != "na") && (point != "-"))
                hash_obj[:adj_points] << point.to_f
                hash_obj[:final_rate] << point.to_f
                hash_obj[:adj_primary_key] << adj.data.keys.first
              end
            end
          end
        end
      end
      if hash_obj[:adj_points].present?
        hash_obj[:final_rate] << hash_obj[:base_rate].to_f
        @result << hash_obj
      else
        hash_obj[:adj_points] = "Adjustment Not Present"
        hash_obj[:adj_primary_key] = "Adjustment Not Present"
        hash_obj[:final_rate] << hash_obj[:base_rate].to_f
        @result << hash_obj
      end


      hash_obj = {
      :bank_name => "",
      :loan_category => "",
      :program_category => "",
      :program_name => "",
      :base_rate => 0.0,
      :adj_points => [],
      :adj_primary_key => [],
      :final_rate => []
    }

    end
  end

  # render :index
end
