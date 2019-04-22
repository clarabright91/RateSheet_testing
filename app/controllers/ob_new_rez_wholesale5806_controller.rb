class ObNewRezWholesale5806Controller < ApplicationController
  method_names = [:government, :programs, :freddie_fixed_rate, :conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm, :homeready, :homeready_hb]
  before_action :read_sheet, only: method_names + [:index]
  before_action :get_sheet, only: method_names
  # before_action :check_sheet_empty , only: method_names
  before_action :get_program, only: [:single_program, :program_property]

  require 'roo'
  require 'roo-xls'

  def index
    # HardWorker.perform_async(1)
    @banks = Bank.all
    @sheetlist =[]

      @xlsx.sheets.each do |sheet|
        @sheetlist.push(sheet)
        if (sheet == "Cover Zone 1")
          @sheet_name = sheet
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @xlsx.sheet(sheet).each_with_index do |row, index|
            current_row = index+1
            if row.include?("Mortgagee Clause (Wholesale)")
              address_index = row.find_index("Mortgagee Clause (Wholesale)")
              @address_a = []
              (1..3).each do |n|
                @address_a << @xlsx.sheet(sheet).row(current_row+n)[address_index]
                if n == 3
                  @zip = @xlsx.sheet(sheet).row(current_row+n)[address_index].split.last
                  @state_code = @xlsx.sheet(sheet).row(current_row+n)[address_index].split[2]
                end
              end
            end
            if (row.include?("Phone") && row.include?("General Contacts"))
              phone_index = row.find_index(headers[0])
              general_contacts_index = row.find_index(headers[1])
              c_row = @xlsx.sheet(sheet).row(current_row+1)
              @name = c_row[general_contacts_index]
              @phone = c_row[phone_index]
            end
          end

          @bank = Bank.find_or_create_by(name: @name)
          @bank.update(phone: @phone, address1: @address_a.join, state_code: @state_code, zip: @zip)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
  end

  def cover_zone_1
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def heloc
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def smartseries
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def jumbo_series_c
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def government
    @programs_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Government")
        @sheet_name = sheet
        @program_ids = []
        @sheet = sheet
        @credit_hash = {}
        @loan_hash = {}
        @hb_hash = {}
        @bpc_loan_hash = {}
        @govt_hash = {}
        @second_hash = {}
        @spe_hash = {}
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        new_key = ''
        new_val = ''
        c_val = ''
        (1..95).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program_ids << @program.id
                  @program.adjustments.destroy_all
                  p_name = @title + " " + sheet
                  # rate arm
                  if @title.include?("10yr") || @title.include?("10 Yr")
                    term = 10
                  elsif @title.include?("15yr") || @title.include?("15 Yr")
                    term = 15
                  elsif @title.include?("20yr") || @title.include?("20 Yr")
                    term = 20
                  elsif @title.include?("25yr") || @title.include?("25 Yr")
                    term = 25
                  elsif @title.include?("30yr") || @title.include?("30 Yr")
                    term = 30
                  end

                  if p_name.include?("Fixed")
                    loan_type = "Fixed"
                  elsif p_name.include?("ARM")
                    loan_type = "ARM"
                    arm_benchmark = "LIBOR"
                    arm_margin = 0
                  elsif p_name.include?("Floating")
                    loan_type = "Floating"
                  elsif p_name.include?("Variable")
                    loan_type = "Variable"
                  else
                    loan_type = "Fixed"
                  end

                  # rate arm
                  if @title.include?("3-1 ARM") || @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                    arm_basic = @title.scan(/\d+/)[0].to_i
                  end

                  freddie_mac = false
                  if p_name.downcase.include?("freddie mac")
                    freddie_mac = true
                  end

                  conforming = false
                  if p_name.downcase.include?("conforming") 
                    conforming = true
                  end

                  fannie_mae = false
                  if p_name.downcase.include?("Fannie Mae")
                    fannie_mae = true
                  end
                  # Arm Advanced
                  if @title.downcase.include?("arm")
                    arm_advanced = @title.split("ARM").last.tr('A-Za-z ()', '')
                  end
                  # High Balance
                  if p_name.include?("High Balance")
                    loan_size = "High-Balance"
                  else
                    loan_size = "Conforming"
                  end
                  # Fha, va, usda
                  if p_name.downcase.include?("fha")
                    fha = true
                  end
                  if p_name.downcase.include?("va")
                    va = true
                  end
                  if p_name.downcase.include?("usda")
                    usda = true
                  end
                  # LoanPurpose
                  if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                    loan_purpose = "Refinance"
                  else
                    loan_purpose = "Purchase"
                  end
                  # lp and du
                  if p_name.downcase.include?('du ')
                    du = true
                  end
                  if p_name.downcase.include?('lp ')
                    lp = true
                  end
                  @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fha: fha, va: va, usda: usda, fannie_mae: fannie_mae, loan_size: loan_size, loan_category: @sheet_name, arm_basic: arm_basic, arm_advanced: arm_advanced, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                  @block_hash = {}
                  key = ''
                  (1..23).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
                        end
                        @data << value
                      end
                    end
                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (110..136).each do |r|
          row = sheet_data.row(r)
          # @key_data = sheet_data.row(40)
          if (row.compact.count >= 1)
            (0..18).each do |max_column|
              cc = max_column
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @credit_hash["FICO"] = {}
                    @loan_hash["LoanAmount/LoanPurpose"] = {}
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"] = {}
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"] = {}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]={}
                    @spe_hash["LoanType"] = {}
                    @second_hash["LoanType/LockDay"]={}
                  end
                  if r >= 112 && r <= 120 && cc == 5
                    new_key = get_value value
                    new_val = sheet_data.cell(r,cc+4)
                    @credit_hash["FICO"][new_key] = new_val
                  end
                  if r >= 123 && r <= 127 && cc == 5
                    if value.downcase.include?("conforming")
                      new_key = "300000-Inf"
                    else
                      new_key = get_value value
                    end
                    new_val = sheet_data.cell(r,cc+4)
                    c_val = sheet_data.cell(r,cc+5)
                    @loan_hash["LoanAmount/LoanPurpose"][new_key] = {}
                    @loan_hash["LoanAmount/LoanPurpose"][new_key]["Purchase"] = new_val
                    @loan_hash["LoanAmount/LoanPurpose"][new_key]["Refinance"] = c_val
                  end
                  if r == 128 && cc == 5
                    new_val = sheet_data.cell(r,cc+4)
                    @hb_hash["High-Balance"] = {}
                    @hb_hash["High-Balance"] = new_val
                  end
                  if r >= 129 && r <= 133 && cc == 5
                    new_key = get_value value
                    new_val = sheet_data.cell(r,cc+4)
                    c_val = sheet_data.cell(r,cc+5)
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"][new_key] = {}
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"][new_key]["Purchase"] = new_val
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"][new_key]["Refinance"] = c_val
                  end
                  if r == 136 && cc == 6
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]={}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]["IRRRL"]={}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]["IRRRL"]["true"]={}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]["IRRRL"]["true"]["true"]=value
                  end
                  if r >= 112 && r <= 125 && cc == 12
                    if value == "30, 45 & 60 Day Lock Purchase Special"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["LoanType/LockDay"]["Purchase"]={}
                      @second_hash["LoanType/LockDay"]["Purchase"][30]=new_val
                      @second_hash["LoanType/LockDay"]["Purchase"][45]=new_val
                      @second_hash["LoanType/LockDay"]["Purchase"][60]=new_val
                    end
                    if value == "FHA Refinances"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/LoanPurpose"]={}
                      @second_hash["FHA/LoanPurpose"]["true"]={}
                      @second_hash["FHA/LoanPurpose"]["true"]["Refinance"]=new_val
                    end
                    if value == "FHA/VA ARM <660"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/LoanType/VA/FICO"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]["ARM"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]["ARM"]["true"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]["ARM"]["true"]["0-660"]=new_val
                    end
                    if value == "90 Day Lock (FRM & Purch Only)"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["RateType/LoanPurpose/LockDay"]={}
                      @second_hash["RateType/LoanPurpose/LockDay"]["Fixed"]={}
                      @second_hash["RateType/LoanPurpose/LockDay"]["Fixed"]["Purchase"]={}
                      @second_hash["RateType/LoanPurpose/LockDay"]["Fixed"]["Purchase"]["90"]=new_val
                    end
                    if value == "VA Cashout >95 LTV"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["RefinanceOption/VA/LTV"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]["true"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]["true"]["LTV"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]["true"]["LTV"]["0-95"]=new_val
                    end
                    if value == "VA - Refinance Credit Score ≥ 620"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["LoanType/VA/FICO"]={}
                      @second_hash["LoanType/VA/FICO"]["Refinance"]={}
                      @second_hash["LoanType/VA/FICO"]["Refinance"]["true"]={}
                      @second_hash["LoanType/VA/FICO"]["Refinance"]["true"]["0-620"]=new_val
                    end
                    if value == "VA - All Loan Purposes - Credit Score < 620"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["VA/FICO"]={}
                      @second_hash["VA/FICO"]["true"]={}
                      @second_hash["VA/FICO"]["true"]["0-620"]=new_val
                    end
                    if value == "VA - IRRRL - Investment Property"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["VA/LoanType/RefinanceOption"]={}
                      @second_hash["VA/LoanType/RefinanceOption"]["true"]={}
                      @second_hash["VA/LoanType/RefinanceOption"]["true"]["Refinance"]={}
                      @second_hash["VA/LoanType/RefinanceOption"]["true"]["Refinance"]["IRRRL"]=new_val
                    end
                    if value == "Manufactured Home (FHA Only)"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/PropertyType"]={}
                      @second_hash["FHA/PropertyType"]["true"]={}
                      @second_hash["FHA/PropertyType"]["true"]["Manufactured Home"]=new_val
                    end
                    if value == "High Balance - 15 Yr Term\n(Adjusting 15 Yr Conforming Pricing - FHA/VA ONLY"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/VA/LoanSize/Term"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]["High-Balance"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]["High-Balance"]["15"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]["High-Balance"]["15"]=new_val
                    end
                    if value == "Margin on all Government ARMs"
                      new_val = sheet_data.cell(r,cc+6)
                      new_val = get_value new_val
                      @second_hash["Margin"]={}
                      @second_hash["Margin"]=new_val
                    end
                  end
                  if r >= 132 && r <= 133 && cc == 17
                    new_val = sheet_data.cell(r,cc+1)
                    @spe_hash["LoanType"]["fixed"] = new_val if value == "Fixed"
                    @spe_hash["LoanType"]["ARM"] = new_val if value == "ARM"
                  end

                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@spe_hash,@hb_hash,@credit_hash,@bpc_loan_hash,@govt_hash,@loan_hash,@second_hash]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(sheet)
      end
    end

    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def freddie_fixed_rate
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Freddie Fixed Rate")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @sheet = sheet
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        main_key = ''
        (1..118).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Freddie Mac 10yr Super Conforming"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                #title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                #term
                term = nil
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = 10
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = 15
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = 20
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = 25
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = 30
                end

                # interest type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # conforming
                if p_name.downcase.include?("super conforming")
                  loan_size = "Super Conforming"
                  conforming = true
                else
                  loan_size = "Conforming"
                end

                # freddie_mac
                if p_name.include?("Freddie Mac")
                  freddie_mac = true
                end

                # fannie_mae
                if p_name.downcase.include?("fannie mae")
                  fannie_mae = true
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, loan_category: @sheet_name, fannie_mae: fannie_mae,loan_size: loan_size, loan_purpose: loan_purpose, du: du, lp: lp,arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''

                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      # first_row[c_i]
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (120..175).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(122)
          @lpmi = sheet_data.row(137)
          @fico = sheet_data.row(158)
          @unit_data = sheet_data.row(155)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 123 && r <= 130 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 123 && r <= 130 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 132 && r <= 135 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 132 && r <= 135 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 138 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 138 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 139 && r <= 142 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 139 && r <= 142 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r == 143 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"] = {}
                  end
                  if r == 143 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = value
                  end
                  if r >= 145 && r <= 148 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 145 && r <= 148 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 150 && r <= 153 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 150 && r <= 153 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 156 && r <= 157 && cc == 6
                    primary_key = value.split('s').first
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 156 && r <= 157 && cc >= 9 && cc <= 11
                    unit_data = get_value @unit_data[cc-2]
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = {}
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = value
                  end
                  if r >= 159 && r <= 162 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 159 && r <= 162 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 159 && r <= 162 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 163 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 164 && cc == 11
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 165 && cc == 11
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r == 166 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 167 && r <= 169 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 156 && r <= 162 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 156 && r <= 162 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 156 && r <= 162 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 162 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 162 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 162 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 163 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 164 && cc == 17
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Purchase"]["Rate and Term"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Refinance"]["Rate and Term"] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Purchase"]["Rate and Term"] = new_val
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Refinance"]["Rate and Term"] = new_val
                  end
                  if r == 165 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["Super Conforming"]["Cash Out"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["Super Conforming"]["Cash Out"] = value
                  end
                  if r == 172 && cc == 17
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r == 173 && cc == 17
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end
    # create adjustment for each program
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def conforming_fixed_rate
    program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Fixed Rate")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..118).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Fannie Mae 10yr High Balance"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                #title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                #term
                term = nil
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = @title.scan(/\d+/)[0]
                end

                # interest type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # conforming
                if p_name.downcase.include?("conforming")
                  conforming = true
                end

                if p_name.downcase.include?("freddie mac")
                  freddie_mac = true
                end

                # fannie_mae
                fannie_mae = false
                if p_name.downcase.include?("fannie mae") 
                  fannie_mae = true
                end

                # High Balance
                if p_name.include?("High Balance")
                  loan_size = "High-Balance"
                else
                  loan_size = "Conforming"
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_ids << @program.id

                @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae,loan_size: loan_size, loan_category: @sheet_name, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (120..173).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(122)
          @lpmi = sheet_data.row(137)
          @fico = sheet_data.row(155)
          @unit_data = sheet_data.row(155)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 123 && r <= 130 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 123 && r <= 130 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 133 && r <= 135 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 133 && r <= 135 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 138 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 138 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 139 && r <= 142 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 139 && r <= 142 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 145 && r <= 148 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 145 && r <= 148 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 150 && r <= 153 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 150 && r <= 153 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 156 && r <= 161 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 156 && r <= 161 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 156 && r <= 161 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 162 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = value
                  end
                  if r == 163 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r == 164 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 165 && cc == 11
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 166 && cc == 11
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r >= 167 && r <= 169 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 156 && r <= 162 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 156 && r <= 162 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 156 && r <= 162 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 162 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 162 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 162 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 163 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 169 && cc == 17
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r == 170 && cc == 17
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = value
                  end
                  if r == 171 && cc == 17
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Purchase"]["Rate and Term"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Refinance"]["Rate and Term"] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Purchase"]["Rate and Term"] = new_val
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Refinance"]["Rate and Term"] = new_val
                  end
                  if r == 172 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["High-Balance"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = value
                  end               
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end

    # create adjustment for each program
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def home_possible
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Home Possible")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..76).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                #title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                #term
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = 10
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = 15
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = 20
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = 25
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = 30
                end

                # rate type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # Arm Basic
                if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end
                # Arm Advanced
                if @title.downcase.include?("arm")
                  arm_advanced = @title.downcase.split("arm").last.tr('A-Za-z() ','')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end
                # freddie_mac
                freddie_mac = false
                if p_name.include?("Freddie Mac") || p_name.downcase.include?("fhlmc")
                  freddie_mac = true
                end

                # fannie_mae
                fannie_mae = false
                if p_name.include?("Fannie Mae") || p_name.include?("Freddie Mac Home Ready")
                  fannie_mae = true
                end
                # freddie_mac_product
                if p_name.downcase.include?("home possible")
                  freddie_mac_product = "Home Possible"
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # loan_size
                if p_name.downcase.include?('high balance')
                  loan_size = "High Balance"
                else
                  loan_size = "Conforming"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                @program.adjustments.destroy_all
                @program.update(term: term,loan_type: loan_type,freddie_mac: freddie_mac, fannie_mae: fannie_mae, loan_category: @sheet_name,arm_basic: arm_basic, freddie_mac_product: freddie_mac_product, arm_advanced: arm_advanced, loan_purpose: loan_purpose, du: du, lp: lp,loan_size: loan_size, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (78..134).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(80)
          @lpmi = sheet_data.row(92)
          @fico = sheet_data.row(114)
          @unit = sheet_data.row(123)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                  end
                  if value == "LPMI Adjustments Applied after Cap"
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  if value == "Adjustments Applied after Cap"
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 81 && r <= 88 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 81 && r <= 88 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r == 93 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 93 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 94 && r <= 95 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 94 && r <= 95 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 97 && r <= 100 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 97 && r <= 100 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 102 && r <= 105 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 102 && r <= 105 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r == 107 && cc == 4
                    @property_hash["PropertyType/LTV"]["Manufactured Home"] = {}
                  end
                  if r == 107 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["PropertyType/LTV"]["Manufactured Home"][ltv_key] = {}
                    @property_hash["PropertyType/LTV"]["Manufactured Home"][ltv_key] = value
                  end
                  if r >= 109 && r <= 112 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"][primary_key] = {}
                  end
                  if r >= 109 && r <= 112 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"][primary_key][ltv_key] = value
                  end
                  if r >= 115 && r <= 118 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 115 && r <= 118 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 115 && r <= 118 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 120 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 121 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 124 && r <= 125 && cc == 6
                    primary_key = value.split('s').first
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 124 && r <= 125 && cc >= 9 && cc <= 10
                    unit_data = get_value @unit[cc-2]
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = {}
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = value
                  end
                  if r == 127 && cc == 8
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r >= 116 && r <= 122 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 116 && r <= 122 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 116 && r <= 122 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 122 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 122 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 122 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 123 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 124 && cc == 19
                    @loan_amount["MiscAdjuster"] = {}
                    @loan_amount["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @loan_amount["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 125 && cc == 19
                    @loan_amount["PropertyType"] = {}
                    @loan_amount["PropertyType"]["Manufactured Home"] = {}
                    @loan_amount["PropertyType"]["Manufactured Home"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  # def lp_open_acces_arms
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "LP Open Acces ARMs")
  #      sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       @unit_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       misc_adj_key = ''
  #       term_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       misc_key = ''
  #       fixed_key = ''
  #       sub_data = ''
  #       key = ''
  #       @sheet = sheet
  #       (1..35).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
  #           rr = r + 1
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # interest type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #                # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustments
  #       (37..71).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(39)
  #         @sub_data = sheet_data.row(47)
  #         @unit_data = sheet_data.row(56)
  #         if row.compact.count >= 1
  #           (0..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All LP Open Access ARMs" || value == "Subordinate Financing" || value == "Number Of Units" || value == "Loan Size Adjustments"
  #                   secondry_key = value
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All LP Open Access ARMs
  #                 if r >= 40 && r<= 45 && cc == 8# && cc <= 19 && cc != 15
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 40 && r<= 45 && cc > 8 && cc != 15 && cc <= 19
  #                   fixed_key = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
  #                 end

  #                 # Subordinate Financing Adjustments
  #                 if r >= 48 && r <= 54 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 48 && r <= 54 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 48 && r<= 54 && cc >= 9 && cc <= 10
  #                   sub_data = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
  #                 end

  #                 # Number Of Units Adjustments
  #                 if r >= 57 && r <= 58 && cc == 3
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 57 && r <= 58 && cc > 3 && cc <= 7
  #                   sub_data = get_value @unit_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][sub_data] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][sub_data] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 61 && r <= 67 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 61 && r <= 67 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r >= 69 && r <= 71 && cc == 3
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 69 && r <= 71 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if  value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   key = value
  #                   @adjustment_hash[key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 47 && r <= 57 && cc == 15
  #                   if value.include?("Condo")
  #                     misc_key = "Condo=>75.01=>15.01"
  #                   else
  #                     misc_key = value
  #                   end
  #                   @adjustment_hash[key][misc_key] = {}
  #                 end
  #                 if r >= 47 && r <= 57 && cc == 19
  #                   @adjustment_hash[key][misc_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r >= 62 && r <= 65 && cc == 16
  #                   misc_key = value
  #                   @adjustment_hash[key][misc_key] = {}
  #                 end
  #                 if r >= 62 && r <= 65 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[key][misc_key][term_key] = {}
  #                 end
  #                 if r >= 62 && r <= 65 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[key][misc_key][term_key][ltv_key] = {}
  #                 end
  #                 if r >= 62 && r <= 65 && cc == 19
  #                   @adjustment_hash[key][misc_key][term_key][ltv_key] = value
  #                 end
  #                 if r >= 67 && r <= 68 && cc == 12
  #                   misc_key = value
  #                   @adjustment_hash[key][misc_key] = {}
  #                 end
  #                 if r >= 67 && r <= 68 && cc == 16
  #                   @adjustment_hash[key][misc_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @sheet)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

  # def lp_open_access_105
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "LP Open Access_105")
  #       sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       term_key = ''
  #       caps_key = ''
  #       max_key = ''
  #       fixed_key = ''
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access 10yr Fixed >125 LTV"))
  #           rr = r + 1
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # interest type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # interest sub type
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustment
  #       (63..86).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(65)
  #         @sub_data = sheet_data.row(68)
  #         if row.compact.count >= 1
  #           (0..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
  #                   secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Subordinate Financing"
  #                   secondry_key = "FinancingType/LTV/CLTV/FICO"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Number Of Units"
  #                   secondry_key = "PropertyType/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == 'Loan Size Adjustments'
  #                   secondry_key = "Loan Size Adjustments"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All Fixed Conforming Adjustments
  #                 if r == 66 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r == 66 && cc > 6 && cc <= 19 && cc != 15
  #                   fixed_key = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
  #                 end

  #                 # Subordinate Financing
  #                 if r == 69 && cc == 5
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r == 69 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r == 69 && cc >= 9 && cc <= 10
  #                   fixed_key = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = value
  #                 end

  #                 # Number Of Units
  #                 if r >= 72 && r <= 73 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 72 && r <= 73 && cc == 5
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 76 && r <= 82 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 76 && r <= 82 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r >= 84 && r <= 86 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 84 && r <= 86 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)
  #               if value.present?
  #                 if  value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   @key = value
  #                   @adjustment_hash[primary_key][@key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 68 && r <= 72 && cc == 15
  #                   if value.include?("Condo")
  #                     cltv_key = "Condo=>105=>15.01"
  #                   else
  #                     cltv_key = value
  #                   end
  #                   @adjustment_hash[primary_key][@key][cltv_key] = {}
  #                 end
  #                 if r >= 68 && r <= 72 && cc == 19
  #                   @adjustment_hash[primary_key][@key][cltv_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r > 76 && r <= 79 && cc == 16
  #                   caps_key = value
  #                   @adjustment_hash[primary_key][@key][caps_key] = {}
  #                 end
  #                 if r > 76 && r <= 79 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
  #                 end
  #                 if r > 76 && r <= 79 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
  #                 end
  #                 if r > 76 && r <= 79 && cc == 19
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r == 82 && cc == 12
  #                   max_key = value
  #                   @adjustment_hash[primary_key][max_key] = {}
  #                 end
  #                 if r == 82 && cc == 16
  #                   @adjustment_hash[primary_key][max_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @program_ids)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

  def jumbo_series_d
    @adjustment_hash = {}
    @property_hash = {}
    @state = {}
    primary_key = ''
    ltv_key = ''
    secondry_key = ''
    @programs_ids =[]
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_D")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        (1..22).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (2 / 8 / 14)
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                program_heading = @title.split
                term =  program_heading[3]
                loan_type = program_heading[5]
                if p_name.downcase.include?("jumbo")
                  loan_size = "Jumbo"
                else
                  loan_size = "Conforming"
                end
                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids  << @program.id
                @program.update(term: term,loan_type: loan_type, loan_category: @sheet_name,loan_size: loan_size, loan_purpose: loan_purpose, du: du, lp: lp)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {} if key.present?
                    else
                      begin
                        @block_hash[key][15*c_i] = value if key.present? &&value.present?
                      rescue Exception => e
                      end
                    end
                    @data << value
                  end
                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #For Adjustments
        (41..71).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(44)
          if row.count >= 1
            (0..17).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "FICO/LTV Adjustments - Loan Amount ≤ $1MM"
                    @adjustment_hash["LoanAmount/FICO/LTV"] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.downcase}"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["PropertyType/LTV"] = {}
                  end
                  if value == "State Adjustments"
                    @state["State"] = {}
                  end
                  # FICO/LTV Adjustments - Loan Amount ≤ $1MM
                  if r >= 45 && r <= 51 && cc == 3
                    if value.include?(">")
                      primary_key = value.tr('>= ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = value
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"][primary_key] = {}
                  end
                  if r >= 45 && r <= 51 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"][primary_key][ltv_key] = value
                  end
                  # State Adjustments
                  if r >= 45 && r <= 61 && cc == 11
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  if r >= 45 && r <= 61 && cc == 13
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  if r >= 45 && r <= 61 && cc == 16
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  # FICO/LTV Adjustments - Loan Amount > $1MM
                  if r >= 55 && r <= 61 && cc == 3
                    if value.include?(">")
                      primary_key = value.tr('>= ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = value
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.downcase}"][primary_key] = {}
                  end
                  if r >= 55 && r <= 61 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.downcase}"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.downcase}"][primary_key][ltv_key] = value
                  end
                  # Max Price
                  if r == 64 && cc == 11
                    @adjustment_hash["LoanType/Term"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"]["20"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"]["30"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @adjustment_hash["LoanType/Term"]["Fixed"]["20"] = new_val
                    @adjustment_hash["LoanType/Term"]["Fixed"]["30"] = new_val
                  end
                  if r == 65 && cc == 11
                    @adjustment_hash["LoanType/Term"]["Fixed"]["15"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @adjustment_hash["LoanType/Term"]["Fixed"]["15"] = new_val
                  end
                  # Feature Adjustments
                  if r >= 65 && r <= 67 && cc == 2
                    if value == "Investment"
                      primary_key = "Investment Property"
                    else
                      primary_key = value
                    end
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 65 && r <= 67 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = value
                  end
                  if r == 68 && cc == 2
                    @property_hash["RefinanceOption/LTV"] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"] = {}
                  end
                  if r == 68 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = value
                  end
                  if r >= 69 && r <= 70 && cc == 2
                    primary_key = value
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 69 && r <= 70 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = value
                  end
                  if r == 70 && cc == 2
                    @property_hash["MiscAdjuster/LTV"] = {}
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"] = {}
                  end
                  if r == 70 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"][ltv_key] = {}
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"][ltv_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@state]
        create_adjust(adjustment,sheet)
      end
    end
    # create adjustment for each program
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  # def lp_open_access
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "LP Open Access")
    # sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       @unit_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       unit_key = ''
  #       caps_key = ''
  #       term_key = ''
  #       max_key = ''
  #       fixed_key = ''
  #       sub_data = ''
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access Super Conforming 10 Yr Fixed"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               arm_basic = false
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae =false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end

  #       # Adjustment
  #       (63..97).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(65)
  #         @sub_data = sheet_data.row(73)
  #         @unit_data = sheet_data.row(82)
  #         if row.compact.count >= 1
  #           (0..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)

  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
  #                   secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Subordinate Financing"
  #                   secondry_key = "FinancingType/LTV/CLTV/FICO"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Number Of Units"
  #                   secondry_key = "PropertyType/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == 'Loan Size Adjustments'
  #                   secondry_key = "Loan Size Adjustments"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All fixed Adjustment
  #                 if r >= 66 && r <= 71 && cc == 8
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 66 && r <= 71 && cc > 8 && cc <= 19 && cc != 15
  #                   fixed_key = @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
  #                 end

  #                 # Subordinate Adjustment
  #                 if r >= 74 && r <= 80 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 74 && r <= 80 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 74 && r <= 80 && cc >= 9 && cc <= 10
  #                   fixed_key = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = value
  #                 end

  #                 # Number of unit Adjustment
  #                 if r >= 83 && r <= 84 && cc == 3
  #                   unit_key = value
  #                   @adjustment_hash[primary_key][secondry_key][unit_key] = {}
  #                 end
  #                 if r >= 83 && r <= 84 && cc > 3 && cc <= 7
  #                   fixed_key = get_value @unit_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][unit_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][unit_key][fixed_key] = value
  #                 end

  #                 # Loan Size Adjustments
  #                 if r >= 87 && r <= 93 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 87 && r <= 93 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 95 && r <= 97 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 95 && r <= 97 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)
  #               if value.present?
  #                 if  value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   @key = value
  #                   @adjustment_hash[primary_key][@key] = {}
  #                 end

  #                 # Misc Adjustment
  #                 if r >= 73 && r <= 80 && cc == 15
  #                   if value.include?("Condo")
  #                     cltv_key = "Condo=>75.01=>15.01"
  #                   else
  #                     cltv_key = value
  #                   end
  #                   @adjustment_hash[primary_key][@key][cltv_key] = {}
  #                 end
  #                 if r >= 73 && r <= 80 && cc == 19
  #                   @adjustment_hash[primary_key][@key][cltv_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r > 86 && r <= 90 && cc == 16
  #                   caps_key = value
  #                   @adjustment_hash[primary_key][@key][caps_key] = {}
  #                 end
  #                 if r > 86 && r <= 90 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
  #                 end
  #                 if r > 86 && r <= 90 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
  #                 end
  #                 if r > 86 && r <= 90 && cc == 19
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
  #                 end


  #                 if r == 93 && cc == 12
  #                   max_key = value
  #                   @adjustment_hash[primary_key][max_key] = {}
  #                 end
  #                 if r == 93 && cc == 16
  #                   @adjustment_hash[primary_key][max_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @program_ids)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

  def jumbo_series_f
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_F")
        @sheet_name = sheet
        @adjustment_hash = {}
        @refinance_hash = {}
        @loan_amount = {}
        @state = {}
        @property_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        (2..36).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 23)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                # term
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = 10
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = 15
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = 20
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = 25
                elsif @title.include?("30yr") || @title.include?("30 Yr") 
                  term = 30
                end

                if @title.include?("20/25/30 Yr")
                  term = 2030
                elsif @title.include?("10/15 Yr")
                  term = 1015
                end

                # rate type
                if p_name.downcase.include?("fixed")
                  loan_type = "Fixed"
                elsif p_name.downcase.include?("arm")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # rate arm
                if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end

                # Arm Advanced
                if @title.downcase.include?("arm")
                  arm_advanced = @title.downcase.split("arm").last.tr('A-Za-z ','')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end

                # Loan Size
                if p_name.downcase.include?("jumbo")
                  loan_size = "Jumbo"
                else
                  loan_size = "Conforming"
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program.update(term: term,loan_type: loan_type,arm_basic: arm_basic, loan_category: @sheet_name,arm_advanced: arm_advanced, loan_size: loan_size, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash, loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        # Adjustments
        (55..94).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(60)
          @cltv_data2 = sheet_data.row(59)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..25).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Purchase Transactions"
                    @adjustment_hash["LoanPurpose/FICO/LTV"] = {}
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"] = {}
                    @state["State"] = {}
                  end
                  if value == "R/T Refinance Transactions"
                    @refinance_hash["RefinanceOption/FICO/LTV"] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                  end
                  if value == "Loan Amount Adjustments"
                    @loan_amount["LoanAmount/LTV"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["PropertyType/LTV"] = {}
                  end
                  # Loan Amount Adjustments
                  if r >= 60 && r <= 63 && cc == 15
                    if value.include?("≤")
                      ltv_key = "0-"+value.tr('A-Z≤ $ ','')+"000000"
                    else
                      ltv_key = (value.tr('A-Z$ ','').split("-").first.to_f*1000000).to_s+"-"+(value.tr('A-Z$ ','').split("-").last.to_f*1000000).to_s
                    end
                    @loan_amount["LoanAmount/LTV"][ltv_key] = {}
                  end
                  if r >= 60 && r <= 63 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = {}
                    @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = value
                  end
                  # Purchase Transactions Adjustment
                  if r >= 61 && r <= 65 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = get_value value
                    end
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key] = {}
                  end
                  if r >= 61 && r <= 65 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = value
                  end
                  # Feature Adjustments
                  if r >= 68 && r <= 73 && cc == 15
                    primary_key = value
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 68 && r <= 73 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].present? && @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @property_hash["PropertyType/LTV"][primary_key][secondry_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][secondry_key] = value
                  end
                  # R/T Refinance Transactions Adjustment
                  if r >= 69 && r <= 73 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key] = {}
                  end
                  if r >= 69 && r <= 73 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = value
                  end
                  # # C/O Refinance Transactions Adjustment
                  if r >= 77 && r <= 81 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 77 && r <= 81 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = value
                  end
                  # State Adjustments
                  if r == 86 && cc == 3
                    @state["State"]["FL"] = {}
                    @state["State"]["NV"] = {}
                  end
                  if r ==86 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @state["State"]["FL"][secondry_key] = {}
                    @state["State"]["NV"][secondry_key] = {}
                    @state["State"]["FL"][secondry_key] = value
                    @state["State"]["NV"][secondry_key] = value
                  end
                  if r == 87 && cc == 3
                    @state["State"]["CA"] = {}
                  end
                  if r ==87 && cc > 3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @state["State"]["CA"][secondry_key] = {}
                    @state["State"]["CA"][secondry_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@refinance_hash,@loan_amount,@state,@property_hash]
        create_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  # def du_refi_plus_arms
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "Du Refi Plus ARMs")
    # sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       fixed_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       sub_data = ''
  #       misc_key = ''
  #       adj_key = ''
  #       term_key = ''
  #       @sheet = sheet
  #       (1..35).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               arm_basic = false
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               # High Balance
  #               if @title.include?("High Balance")
  #                 jumbo_high_balance = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet, jumbo_high_balance: jumbo_high_balance)
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustments
  #       (37..70).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(39)
  #         @sub_data = sheet_data.row(49)
  #         if row.compact.count >= 1
  #           (3..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All DU Refi Plus Conforming ARMs (All Occupancies)" || value == "Subordinate Financing" || value == "Loan Size Adjustments"
  #                   secondry_key = value
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All du refi plus Adjustment
  #                 if r >= 40 && r <= 47 && cc == 8
  #                   fixed_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key] = {}
  #                 end
  #                 if r >= 40 && r <= 47 && cc >8 && cc <= 19
  #                   fixed_data = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key][fixed_data] = value
  #                 end

  #                 # Subordinate Financing Adjustment
  #                 if r >= 50 && r <= 54 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 50 && r <= 54 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 50 && r <= 54 && cc > 6 && cc <= 10
  #                   sub_data = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 56 && r <= 57 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 56 && r <= 57 && cc == 8
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 60 && r <= 66 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 60 && r <= 66 && cc > 6 && cc <= 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 69 && r <= 70 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 69 && r <= 70 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   misc_key = value
  #                   @adjustment_hash[misc_key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 49 && r <= 58 && cc == 15
  #                   if value.include?("Condo")
  #                     adj_key = "Condo/75"
  #                   else
  #                     adj_key = value
  #                   end
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 49 && r <= 58 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r >= 62 && r <= 64 && cc == 16
  #                   adj_key = value
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 62 && r <= 64 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key] = {}
  #                 end
  #                 if r >= 62 && r <= 64 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = {}
  #                 end
  #                 if r >= 62 && r <= 64 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @sheet)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

  def jumbo_series_h
    @program_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_H")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @sheet = sheet
        @adjustment_hash = {}
        @property_hash = {}
        primary_key = ''
        ltv_key = ''
        secondry_key = ''
        (2..86).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("Jumbo Series H Product and Pricing"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 28)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4 + max_column*6 # (4 / 10 / 16/ 22)
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  p_name = @title + " " + sheet
                  program_heading = @title.split
                  # term
                  term = nil
                  program_heading = @title.split
                  if @title.include?("10yr") || @title.include?("10 Yr")
                    term = 10
                  elsif @title.include?("15yr") || @title.include?("15 Yr")
                    term = 15
                  elsif @title.include?("20yr") || @title.include?("20 Yr")
                    term = 20
                  elsif @title.include?("25yr") || @title.include?("25 Yr")
                    term = 25
                  elsif @title.include?("30yr") || @title.include?("30 Yr")
                    term = 30
                  end

                  # rate type
                  if p_name.downcase.include?("fixed")
                    loan_type = "Fixed"
                  elsif p_name.downcase.include?("arm")
                    loan_type = "ARM"
                    arm_benchmark = "LIBOR"
                    arm_margin = 0
                  elsif p_name.include?("Floating")
                    loan_type = "Floating"
                  elsif p_name.include?("Variable")
                    loan_type = "Variable"
                  else
                    loan_type = "Fixed"
                  end

                  # rate arm
                  if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 Yr ARM") || @title.include?("7/1 Yr ARM") || @title.include?("10/1 Yr ARM")
                    arm_basic = @title.scan(/\d+/)[0].to_i
                  end

                  # Arm Advanced
                  if @title.downcase.include?("arm")
                    arm_advanced = @title.downcase.split("arm").last.tr('A-Za-z- ','')
                    if arm_advanced.include?('/')
                      arm_advanced = arm_advanced.tr('/','-')
                    else
                      arm_advanced
                    end
                  end

                  # freddie_mac
                  freddie_mac = false
                  if p_name.include?("Freddie Mac")
                    freddie_mac = true
                  end

                  # fannie_mae
                  fannie_mae = false
                  if p_name.include?("Fannie Mae") || p_name.include?("Freddie Mac Home Ready")
                    fannie_mae = true
                  end

                  # High Balance
                  if p_name.downcase.include?("jumbo")
                    loan_size = "Jumbo"
                  else
                    loan_size = "Conforming"
                  end

                  # Purchase & Refinance
                  # loan_purpose
                  if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                    loan_purpose = "Refinance"
                  else
                    loan_purpose = "Purchase"
                  end
                  # lp and du
                  if p_name.downcase.include?('du ')
                    du = true
                  end
                  if p_name.downcase.include?('lp ')
                    lp = true
                  end
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program_ids << @program.id
                  @program.update(term: term,loan_type: loan_type,loan_purpose: loan_purpose ,arm_basic: arm_basic, loan_category: @sheet_name,arm_advanced: arm_advanced,loan_size: loan_size, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                  @program.adjustments.destroy_all

                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        if @program.lock_period.length <= 3
                          @program.lock_period << 15*c_i
                          @program.save
                        end
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end

                    if @data.compact.length == 0
                      break # terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil?
                    @block_hash.values.first.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #For Adjustments
        (68..97).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(71)
          @credit_data = sheet_data.row(84)
          if row.count >= 1
            (0..20).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Jumbo Series H - Adjustments"
                    @adjustment_hash["LoanSize/LoanType/State/Term"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/FICO/LTV"] = {}
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"] = {}
                    @property_hash["RefinanceOption/LTV"] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"] = {}
                  end
                  if r >= 72 && r <= 81 && cc == 12
                    primary_key = value
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 72 && r <= 81 && cc >= 13 && cc <= 17
                    ltv_data = @ltv_data[cc-2].tr('A-Za-z ','')
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"][primary_key][ltv_data] = value
                  end
                  if r >= 72 && r <= 81 && cc == 12
                    primary_key = value
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 72 && r <= 81 && cc >= 18 && cc <= 20
                    ltv_data = @ltv_data[cc-2].tr('A-Za-z ','').split('/').first
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r >= 85 && r <= 90 && cc == 12
                    primary_key = get_value value
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"][primary_key] = {}
                  end
                  if r >= 85 && r <= 90 && cc >= 13 && cc <= 17
                    credit_data = get_value @credit_data[cc-2]
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"][primary_key][credit_data] = {}
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"][primary_key][credit_data] = value
                  end
                  if r == 93 && cc == 17
                    @property_hash["LoanAmount"] = {}
                    @property_hash["LoanAmount"]["100000-Inf"] = {}
                    @property_hash["LoanAmount"]["100000-Inf"] = value
                  end
                  if r == 94 && cc == 17
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2nd Home"] = {}
                    @property_hash["PropertyType"]["2nd Home"] = value
                  end
                  if r >= 95 && r <= 97 && cc == 15
                    primary_key = get_value value
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][primary_key] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][primary_key] = new_val
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash]
        create_adjust(adjustment,sheet) 
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  # def du_refi_plus_fixed_rate_105
  #   @program_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "Du Refi Plus Fixed Rate_105")
    # sheet_data = @xlsx.sheet(sheet)
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("DU Refi Plus 10yr Fixed >125 LTV"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               main_key = ''
  #               if @program.term.present?
  #                 main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               else
  #                 main_key = "InterestRate/LockPeriod"
  #               end
  #               @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[main_key][key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[main_key][key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end

  #       #For Adjustments
  #       @xlsx.sheet(sheet).each_with_index do |sheet_row, index|
  #         index = index+ 1
  #         if sheet_row.include?("Loan Level Price Adjustments: See Adjustment Caps")
  #           (index..@xlsx.sheet(sheet).last_row).each do |adj_row|
  #             # First Adjustment
  #             if adj_row == 65
  #               key = ''
  #               key_array = []
  #               rr = adj_row
  #               cc = 3
  #               @occupancy_hash = {}
  #               main_key = "All Occupancies"
  #               @occupancy_hash[main_key] = {}

  #               (0..2).each do |max_row|
  #                 column_count = 0
  #                 rrr = rr + max_row
  #                 row = @xlsx.sheet(sheet).row(rrr)

  #                 if rrr == rr
  #                   row.compact.each do |row_val|
  #                     val = row_val.split
  #                     if val.include?("<")
  #                       key_array << 0
  #                     else
  #                       key_array << row_val.split("-")[0].to_i.round if row_val.include?("-")
  #                       key_array << row_val.split[1].to_i.round if row_val.include?(">")
  #                     end
  #                   end
  #                 end

  #                 (0..16).each do |max_column|
  #                   ccc = cc + max_column
  #                   begin
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc)
  #                     if row.include?("All Occupancies > 15 Yr Terms")
  #                       if value != nil && value.to_s.include?(">") && value != "All Occupancies > 15 Yr Terms" && !value.is_a?(Numeric)
  #                         key = value.gsub(/[^0-9A-Za-z]/, '')
  #                         @occupancy_hash[main_key][key] = {}
  #                       elsif (value != nil) && !value.is_a?(String)
  #                         @occupancy_hash[main_key][key][key_array[column_count]] = value
  #                         column_count = column_count + 1
  #                       end
  #                     end
  #                   rescue Exception => e
  #                     error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                     error_log.save
  #                   end
  #                 end
  #               end
  #               make_adjust(@occupancy_hash, @program_ids)
  #             end

  #             # Second Adjustment(Adjustment Caps)
  #             if adj_row == 86
  #               key_array = ""
  #               rr = adj_row
  #               cc = 16
  #               @adjustment_cap = {}
  #               main_key = "Adjustment Caps"
  #               @adjustment_cap[main_key] = {}
  #               key = ''

  #               (0..4).each do |max_row|
  #                 column_count = 1
  #                 rrr = rr + max_row
  #                 row = @xlsx.sheet(sheet).row(rrr)
  #                 if rrr == 86
  #                   key_array = row.compact
  #                 end

  #                 (0..3).each do |max_column|
  #                   ccc = cc + max_column
  #                   begin
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc)
  #                     if ccc == 16
  #                       key = value if value != nil
  #                       @adjustment_cap[main_key][key] = {} if value != nil
  #                     else
  #                       if !key_array.include?(value)
  #                         @adjustment_cap[main_key][key][key_array[column_count]] = value if value != nil
  #                         column_count = column_count + 1 if value != nil
  #                       end
  #                     end
  #                   rescue Exception => e
  #                     error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #                     error_log.save
  #                   end
  #                 end
  #               end
  #               make_adjust(@adjustment_cap, @program_ids)
  #             end

  #             # Third Adjustment
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Max YSP")
  #               rr = adj_row
  #               cc = 4
  #               begin
  #                 @max_ysp_hash = {}
  #                 main_key = "Max YSP"
  #                 @max_ysp_hash[main_key] = {}
  #                 row = @xlsx.sheet(sheet).row(rr)
  #                 @max_ysp_hash[main_key] = row.compact[5]
  #                 make_adjust(@max_ysp_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end

  #             # Fourth Adjustment (Adjustments Applied after Cap)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Loan Size Adjustments")
  #               rr = adj_row
  #               cc = 6
  #               @loan_size = {}
  #               main_key = "Loan Size / Loan Type"
  #               @loan_size[main_key] = {}

  #               (0..6).each do |max_row|
  #                 @data = []
  #                 rrr = rr + max_row
  #                 ccc = cc
  #                 begin
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)
  #                   if key.present?

  #                     if (key.include?("<"))
  #                       key = 0
  #                     elsif (key.include?("-"))
  #                       key = key.split("-").first.tr("^0-9", '')
  #                     else
  #                       key
  #                     end
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     raise "value is nil at row = #{rrr} and column = #{ccc}" unless value || key
  #                     @loan_size[main_key][key] = value
  #                   end
  #                 rescue Exception => e
  #                   error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                   error_log.save
  #                 end
  #               end
  #               make_adjust(@loan_size, @program_ids)
  #             end

  #             # Fifth Adjustment(Misc Adjusters)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #               rr = adj_row
  #               cc = 15
  #               @cando_hash = {}
  #               main_key = "PropertyType/LTV/Term"
  #               @cando_hash[main_key] = {}

  #               (0..6).each do |max_row|
  #                 @data = []
  #                 rrr = rr + max_row
  #                 ccc = cc
  #                 begin
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?("Condo")
  #                     val = key.split
  #                     key1 = "Condo"
  #                     key2 = val[1].gsub(/[^0-9A-Za-z]/, '')
  #                     key3 = val[3].gsub(/[^0-9A-Za-z]/, '').split("yr")[0]
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @cando_hash[main_key][key1] = {}
  #                     @cando_hash[main_key][key1][key2] = {}
  #                     @cando_hash[main_key][key1][key2][key3] = value
  #                   end

  #                   if key == "Manufactured Home"
  #                     key1 = "Manufactured Home"
  #                     key2 = 0
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @cando_hash[main_key][key1] = {}
  #                     @cando_hash[main_key][key1][key2] = {}
  #                     @cando_hash[main_key][key1][key2] = value
  #                   end
  #                 rescue Exception => e
  #                   error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                   error_log.save
  #                 end
  #               end
  #               make_adjust(@cando_hash, @program_ids)
  #             end

  #             # Sixth Adjustment(Misc Adjusters (2-4 Units))
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #                 rr = adj_row
  #                 cc = 15
  #               begin
  #                 @unit_hash = {}
  #                 main_key = "PropertyType/LTV"
  #                 @unit_hash[main_key] = {}

  #                 rrr = rr + 1
  #                 ccc = cc
  #                 key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                 if key.include?("Units")
  #                   key1 = "2-4 unit"
  #                   value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                   @unit_hash[main_key][key1] = {}
  #                   @unit_hash[main_key][key1] = value
  #                 end
  #                 make_adjust(@unit_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end


  #             # Seventh Adjustment(Misc Adjusters)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #               rr = adj_row
  #               cc = 15
  #               begin
  #                 @data_hash = {}
  #                 main_key = "MiscAdjuster"
  #                 @data_hash[main_key] = {}

  #                 (0..2).each do |max_row|
  #                   rrr = rr + max_row
  #                   ccc = cc
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if !key.include?("Units")
  #                     key1 = key.include?(">") ? key.split(" >")[0] : key
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @data_hash[main_key][key1] = {}
  #                     @data_hash[main_key][key1] = value
  #                   end
  #                 end
  #                 make_adjust(@data_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end

  #             # LTV Adjustment(Misc Adjusters)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #               rr = adj_row
  #               cc = 15
  #               @ltv_hash = {}
  #               main_key = "LTV"
  #               @ltv_hash[main_key] = {}

  #               (0..6).each do |max_row|
  #                 rrr = rr + max_row
  #                 ccc = cc
  #                 begin
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?("LTV") && !key.include?("Condo")
  #                     key1 = key.split[1].to_i.round
  #                     key2 = key.include?("<") ? 0 : 30
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @ltv_hash[main_key][key1] = {} if @ltv_hash[main_key] == {}
  #                     @ltv_hash[main_key][key1][key2] = {}
  #                     @ltv_hash[main_key][key1][key2] = value
  #                   end
  #                 rescue Exception => e
  #                   error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                   error_log.save
  #                 end
  #               end
  #               make_adjust(@ltv_hash, @program_ids)
  #             end

  #             # CA Escrow Waiver Adjustment
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Expanded Approval **")
  #               rr = adj_row
  #               cc = 3
  #               begin
  #                 @misc_adjuster = {}
  #                 main_key = "MiscAdjuster"
  #                 @misc_adjuster[main_key] = {}

  #                 (0..2).each do |max_row|
  #                   rrr = rr + max_row
  #                   ccc = cc
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?("CA Escrow Waiver") || key.include?("Expanded Approval **")
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+7)
  #                     @misc_adjuster[main_key][key] = {}
  #                     @misc_adjuster[main_key][key] = value
  #                   end
  #                 end
  #                 make_adjust(@misc_adjuster, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end

  #             # Subordinate Financing Adjustment
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Subordinate Financing")
  #               rr = adj_row
  #               cc = 6
  #               begin
  #                 @subordinate_hash = {}
  #                 main_key = "FinancingType/LTV/CLTV/FICO"
  #                 key1 = "Subordinate Financing"

  #                 sub_key1 = row.compact[2].include?("<") ? 0 : row.compact[2].split(" ")[1].to_i
  #                 sub_key2 = row.compact[3].include?(">") ? row.compact[3].split(" ")[1].to_i : row.compact[3].to_i

  #                 @subordinate_hash[main_key] = {}
  #                 @subordinate_hash[main_key][key1] = {}

  #                 (1..2).each do |max_row|
  #                   rrr = rr + max_row
  #                   ccc = cc
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?(">") || key == "ALL"
  #                     key2 = (key.include?(">")) ? key.gsub(/[^0-9A-Za-z]/, '') : key
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+3)
  #                     value1 = @xlsx.sheet(sheet).cell(rrr,ccc+4)

  #                     @subordinate_hash[main_key][key1][key2] ={}
  #                     @subordinate_hash[main_key][key1][key2][sub_key1] = value
  #                     @subordinate_hash[main_key][key1][key2][sub_key2] = value1
  #                   end
  #                 end
  #                 make_adjust(@subordinate_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end
  #           end
  #         end
  #       end
  #     end
  #   end
  #   create_program_association_with_adjustment(@sheet)
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

  def jumbo_series_i
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_I")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @property_hash = {}
        @state = {}
        primary_key = ''
        ltv_key = ''
        secondry_key = ''
        main_key = ''
        @sheet = sheet
        # programs
        (2..32).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  p_name = @title + " " + sheet
                  program_heading = @title.split
                  # term
                    term = nil
                    program_heading = @title.split
                    if @title.include?("10yr") || @title.include?("10 Yr")
                      term = @title.scan(/\d+/)[0]
                    elsif @title.include?("15yr") || @title.include?("15 Yr")
                      term = @title.scan(/\d+/)[0]
                    elsif @title.include?("20yr") || @title.include?("20 Yr")
                      term = @title.scan(/\d+/)[0]
                    elsif @title.include?("25yr") || @title.include?("25 Yr")
                      term = @title.scan(/\d+/)[0]
                    elsif @title.include?("30yr") || @title.include?("30 Yr")
                      term = @title.scan(/\d+/)[0]
                    end

                    # rate type
                    if p_name.include?("Fixed")
                      loan_type = "Fixed"
                    elsif p_name.include?("ARM")
                      loan_type = "ARM"
                      arm_benchmark = "LIBOR"
                      arm_margin = 0
                    elsif p_name.include?("Floating")
                      loan_type = "Floating"
                    elsif p_name.include?("Variable")
                      loan_type = "Variable"
                    else
                      loan_type = "Fixed"
                    end

                    # rate arm
                    arm_basic = nil
                    if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("7/1") || @title.include?("5/1")
                      arm_basic = @title.scan(/\d+/)[0].to_i
                    end

                    arm_advanced = nil
                    if @title.include?("5/2/5")
                      arm_advanced = "5-2-5"
                    end

                    loan_size = nil
                    if p_name.include?("Jumbo")
                      loan_size = "Jumbo"
                    else
                      loan_size = "Conforming"
                    end

                    # loan_purpose
                    if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                      loan_purpose = "Refinance"
                    else
                      loan_purpose = "Purchase"
                    end

                    # lp and du
                    if p_name.downcase.include?('du ')
                      du = true
                    end
                    if p_name.downcase.include?('lp ')
                      lp = true
                    end

                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program_ids << @program.id
                  @program.update(term: term,loan_type: loan_type ,arm_basic: arm_basic, loan_category: @sheet_name, loan_size: loan_size, arm_advanced: arm_advanced, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end

                    if @data.compact.length == 0
                      break # terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil?
                    @block_hash.values.first.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        #For Adjustments
        (34..73).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(40)
          if row.count >= 1
            (0..19).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Jumbo Series I Adjustments"
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"] = {}
                    @state["LoanSize/LoanType/Term"] = {}
                    @state["LoanSize/LoanType/Term"]["Jumbo"] = {}
                    @state["LoanSize/LoanType/Term"]["Jumbo"]["Fixed"] = {}
                    @state["LoanSize/LoanType/ArmBasic"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"] = {}
                  end
                  if r >= 41 && r <= 45 && cc == 3
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 41 && r <= 45 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = value
                  end
                  if r >= 41 && r <= 45 && cc == 12
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 41 && r <= 45 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r >= 50 && r <= 53 && cc == 3
                    primary_key = convert_range value
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 50 && r <= 53 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = value
                  end
                  if r >= 50 && r <= 53 && cc == 12
                    primary_key = convert_range value
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 50 && r <= 53 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r == 58 && cc == 3
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2nd Home"] = {}
                  end
                  if r == 58 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2nd Home"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2nd Home"][ltv_data] = value
                  end
                  if r == 58 && cc == 12
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2nd Home"] = {}
                  end
                  if r == 58 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2nd Home"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2nd Home"][ltv_data] = value
                  end
                  if r == 59 && cc == 3
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"] = {}
                  end
                  if r == 59 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = value
                  end
                  if r == 59 && cc == 12
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"] = {}
                  end
                  if r == 59 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = value
                  end
                  if r == 60 && cc == 3
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"]["Cash Out"] = {}
                  end
                  if r == 60 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"]["Cash Out"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"]["Cash Out"][ltv_data] = value
                  end
                  if r == 60 && cc == 12
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"] = {}
                  end
                  if r == 60 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = value
                  end
                  if r == 61 && cc == 3
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2-4 Unit"] = {}
                  end
                  if r == 61 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2-4 Unit"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2-4 Unit"][ltv_data] = value
                  end
                  if r == 61 && cc == 12
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2-4 Unit"] = {}
                  end
                  if r == 61 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2-4 Unit"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2-4 Unit"][ltv_data] = value
                  end
                  if r == 62 && cc == 3
                    @property_hash["LoanSize/LoanType/DTI/LTV"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"]["40%"] = {}
                  end
                  if r == 62 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"]["40%"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"]["40%"][ltv_data] = value
                  end
                  if r == 62 && cc == 12
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"]["40%"] = {}
                  end
                  if r == 62 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"]["40%"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"]["40%"][ltv_data] = value
                  end
                  if r == 66 && cc == 7
                    @state["MiscAdjuster/State"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["CA"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NC"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["DC"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NY"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["CA"] = value
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NC"] = value
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["DC"] = value
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NY"] = value
                  end
                  if r == 66 && cc ==17
                    @state["LoanSize/LoanPurpose/LoanType/Term"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"]["Fixed"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"]["Fixed"]["30"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"]["Fixed"]["30"] = value
                  end
                  if r >= 72 && r <= 73 && cc == 3
                    primary_key = value.tr('A-Za-z ','')
                    @state["LoanSize/LoanType/Term"]["Jumbo"]["Fixed"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @state["LoanSize/LoanType/Term"]["Jumbo"]["Fixed"][primary_key] = new_val
                  end
                  if r == 72 && cc == 17
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["7"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["10"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["7"] = value
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["10"] = value
                  end
                  if r == 73 && cc == 17
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["5"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["5"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@state]
        create_adjust(adjustment,sheet)        
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  # def du_refi_plus_fixed_rate
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "Du Refi Plus Fixed Rate")
    # sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       sub_data = ''
  #       primary_key = ''
  #       secondry_key = ''
  #       fixed_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       misc_key = ''
  #       adj_key = ''
  #       term_key = ''
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("DU Refi Plus 10yr Fixed High Balance"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # High Balance
  #               jumbo_high_balance = false
  #               if @title.include?("High Balance")
  #                 jumbo_high_balance = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming, arm_basic: arm_basic, loan_category: sheet, jumbo_high_balance: jumbo_high_balance)
  #               @program.adjustments.destroy_all
  #               @block_hash = {}
  #               key = ''
  #               main_key = ''
  #               if @program.term.present?
  #                 main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               else
  #                 main_key = "InterestRate/LockPeriod"
  #               end
  #               @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[main_key][key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[main_key][key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustments
  #       (63..94).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(65)
  #         @sub_data = sheet_data.row(75)
  #         if row.compact.count >= 1
  #           (3..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if (r == 65 && cc == 3)
  #                   secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Subordinate Financing"
  #                   secondry_key = "FinancingType/LTV/CLTV/FICO"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Loan Size Adjustments"
  #                   secondry_key = "Loan Size Adjustments"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All Fixed Confoming Adjustment
  #                 if r >= 66 && r <= 73 && cc == 8
  #                   fixed_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key] = {}
  #                 end
  #                 if r >= 66 && r <= 73 && cc >8 && cc <= 19
  #                   fixed_data = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key][fixed_data] = value
  #                 end

  #                 # Subordinate Financing Adjustment
  #                 if r >= 76 && r <= 80 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 76 && r <= 80 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 76 && r <= 80 && cc > 6 && cc <= 10
  #                   sub_data = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 83 && r <= 89 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 83 && r <= 89 && cc > 6 && cc <= 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 92 && r <= 94 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 92 && r <= 94 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   misc_key = value
  #                   @adjustment_hash[misc_key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 75 && r <= 83 && cc == 15
  #                   if value.include?("Condo")
  #                     adj_key = "Condo/75/15"
  #                   else
  #                     adj_key = value
  #                   end
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 75 && r <= 83 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r == 85 && cc == 13
  #                   adj_key = value
  #                   @adjustment_hash[adj_key] = {}
  #                 end
  #                 if r == 85 && cc == 17
  #                   @adjustment_hash[adj_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r >= 89 && r <= 93 && cc == 16
  #                   adj_key = value
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 89 && r <= 93 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key] = {}
  #                 end
  #                 if r >= 89 && r <= 93 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = {}
  #                 end
  #                 if r >= 89 && r <= 93 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @program_ids)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

  def jumbo_series_jqm
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_JQM")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @refinance_hash = {}
        @loan_amount = {}
        @state = {}
        @property_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        @sheet = sheet
        (2..60).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                if @title.present?
                  program_heading = @title.split
                  # term
                  term = nil
                  if @title.include?("10yr") || @title.include?("10 Yr")
                    term = 10
                  elsif @title.include?("15yr") || @title.include?("15 Yr")
                    term = 15
                  elsif @title.include?("20yr") || @title.include?("20 Yr")
                    term = 20
                  elsif @title.include?("25yr") || @title.include?("25 Yr")
                    term = 25
                  elsif @title.include?("30yr") || @title.include?("30 Yr")
                    term = 30
                  end
                  if @title.include?("20/25/30 Yr")
                    term = 2030
                  elsif @title.include?("10/15 Yr")
                    term = 1015
                  end
                  # rate type
                  if p_name.include?("Fixed")
                    loan_type = "Fixed"
                  elsif p_name.include?("ARM")
                    loan_type = "ARM"
                    arm_benchmark = "LIBOR"
                    arm_margin = 0
                  elsif p_name.include?("Floating")
                    loan_type = "Floating"
                  elsif p_name.include?("Variable")
                    loan_type = "Variable"
                  else
                    loan_type = "Fixed"
                  end

                  # Loan Size
                  if p_name.downcase.include?("jumbo")
                    loan_size = "Jumbo"
                  else
                    loan_size = "Conforming"
                  end
                  # Arm basic
                  if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 LIBOR ARM") || @title.include?("7/1 LIBOR ARM") || @title.include?("10/1 LIBOR ARM")
                    arm_basic = @title.scan(/\d+/)[0].to_i
                  end
                  # Arm Advanced
                  if @title.downcase.include?("arm")
                    arm_advanced = @title.downcase.split("arm").last.tr("A-Za-z- ","")
                    if arm_advanced.include?('/')
                      arm_advanced = arm_advanced.tr('/','-')
                    else
                      arm_advanced
                    end
                  end
                  # loan_purpose
                  if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                    loan_purpose = "Refinance"
                  else
                    loan_purpose = "Purchase"
                  end
                  # lp and du
                  if p_name.downcase.include?('du ')
                    du = true
                  end
                  if p_name.downcase.include?('lp ')
                    lp = true
                  end
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program_ids << @program.id
                  @program.update(term: term,loan_type: loan_type,arm_basic: arm_basic, loan_category: @sheet_name,loan_size: loan_size,arm_advanced: arm_advanced, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end

                    if @data.compact.length == 0
                      break #terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil?
                    @block_hash.values.first.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (64..99).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(67)
          @cltv_data2 = sheet_data.row(66)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..25).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Purchase Transactions"
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"] = {}
                    @state["LoanSize/State/LTV"] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"] = {}
                  end
                  if value == "R/T Refinance Transactions"
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"] = {}
                  end
                  if value == "Loan Amount Adjustments"
                    @loan_amount["LoanSize/LoanAmount/LTV"] = {}
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["LoanSize/PropertyType/LTV"] = {}
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"] = {}
                  end
                  # Loan Amount Adjustments
                  if r >= 67 && r <= 70 && cc == 15
                    if value.include?("≤")
                      ltv_key = "0-"+value.tr('A-Z≤ $ ','')+"000000"
                    else
                      ltv_key = (value.tr('A-Z$ ','').split("-").first.to_f*1000000).to_s+"-"+(value.tr('A-Z$ ','').split("-").last.to_f*1000000).to_s
                    end
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"][ltv_key] = {}
                  end
                  if r >= 67 && r <= 70 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"][ltv_key][secondry_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"][ltv_key][secondry_key] = value
                  end
                  # Purchase Transactions Adjustment
                  if r >= 68 && r <= 74 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = get_value value
                    end
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"][primary_key] = {}
                  end
                  if r >= 68 && r <= 74 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"][primary_key][secondry_key] = value
                  end
                  # Feature Adjustments
                  if r >= 75 && r <= 78 && cc == 15
                    if value == "Investment"
                      primary_key = "Investment Property"
                    else
                      primary_key = value
                    end
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"][primary_key] = {}
                  end
                  if r >= 75 && r <= 78 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"][primary_key][secondry_key] = {}
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"][primary_key][secondry_key] = value
                  end
                  if r == 79 && cc == 15
                    @property_hash["LoanSize/MiscAdjuster/LTV"] = {}
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"]["Escrow Waiver"] = {}
                  end
                  if r == 79 && cc  > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"]["Escrow Waiver"][secondry_key] = {}
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"]["Escrow Waiver"][secondry_key] = value
                  end
                  # R/T Refinance Transactions Adjustment
                  if r >= 78 && r <= 84 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"][primary_key] = {}
                  end
                  if r >= 78 && r <= 84 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"][primary_key][secondry_key] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"][primary_key][secondry_key] = value
                  end
                  # # C/O Refinance Transactions Adjustment
                  if r >= 88 && r <= 94 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 88 && r <= 94 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"][primary_key][secondry_key] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"][primary_key][secondry_key] = value
                  end
                  # State Adjustments
                  if r == 99 && cc == 3
                    @state["LoanSize/State/LTV"]["Jumbo"]["FL"] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"]["NV"] = {}
                  end
                  if r ==99 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @state["LoanSize/State/LTV"]["Jumbo"]["FL"][secondry_key] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"]["NV"][secondry_key] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"]["FL"][secondry_key] = value
                    @state["LoanSize/State/LTV"]["Jumbo"]["NV"][secondry_key] = value
                  end
                  if r >= 85 && r <= 87 && cc == 15
                    primary_key = value.tr('A-Za-z ','')
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"][primary_key] = {}
                  end
                  if r >= 85 && r <= 87 && cc >= 16 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"][primary_key][secondry_key] = value
                  end
                  if r >= 88 && r <= 90 && cc == 15        
                    if value == "20 yr Fixed\n(add 30 yr)"
                      primary_key = "20-30"
                    else
                      primary_key = value.tr('A-Za-z ','')
                    end
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 88 && r <= 90 && cc >= 16 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"][primary_key][secondry_key] = value
                  end
                  if r == 96 && cc == 16
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"]["Fixed"] = {}

                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"]["ARM"] = {}
                  end
                  if r == 96 && cc >= 19 && cc <= 20
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','')
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"]["Fixed"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"]["Fixed"][max_key] = value
                  end
                  if r == 96 && cc >= 21 && cc <= 23
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','').split("/").first
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"]["ARM"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"]["ARM"][max_key] = value
                  end
                  if r == 97 && cc == 16
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"]["Fixed"] = {}

                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"]["ARM"] = {}
                  end
                  if r == 97 && cc >= 19 && cc <= 20
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','')
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"]["Fixed"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"]["Fixed"][max_key] = value
                  end
                  if r == 97 && cc >= 21 && cc <= 23
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','').split("/").first
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"]["ARM"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"]["ARM"][max_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@refinance_hash,@loan_amount,@state,@property_hash]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def dream_big
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Dream Big")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @jumbo_adjustment = {}
        @cash_out = {}
        @other_hash = {}
        @program_ids = []
        primary_key = ''
        fixed_key = ''
        ltv_key = ''
        @sheet = sheet
        (1..33).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("Dream Big Jumbo"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              begin
                # title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                # term
                term = nil
                program_heading = @title.split
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = 10
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = 15
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = 20
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = 25
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = 30
                end
                if @title.include?("20/25/30")
                  term = 2030
                end
                # interest type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # Arm basic
                if @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10-1 ARM") || @title.include?("10/1 ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end

                # Arm Advanced
                if @title.downcase.include?("arm")
                  arm_advanced = @title.downcase.split("arm").last.tr('A-Za-z- ','')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end
                # conforming
                if p_name.downcase.include?("conforming")
                  conforming = true
                  loan_size = "Conforming"
                else
                  loan_size = "Conforming"
                end

                # freddie_mac
                if p_name.downcase.include?("freddie mac")
                  freddie_mac = true
                end

                # fannie_mae
                if p_name.downcase.include?("fannie mae")
                  fannie_mae = true
                end
                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: @sheet_name,arm_advanced: arm_advanced, loan_purpose: loan_purpose, du: du, lp: lp, loan_size: loan_size, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (38..62).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(39)
          @ltv_arm_data = sheet_data.row(54)
          if row.compact.count >= 1
            (0..18).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "LTV Based Adjustments for 20/25/30 Yr Fixed Jumbo Products"
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"] = {}
                  end
                  if value == "LTV Based Adjustments for 15 Yr Fixed and All ARM Jumbo Products"
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"] = {}
                  end
                  if r >= 40 && r <= 45 && cc == 3
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"][primary_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"][primary_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"][primary_key] = {}
                  end
                  if r >= 40 && r <= 45 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"][primary_key][ltv_data] = value
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"][primary_key][ltv_data] = value
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"][primary_key][ltv_data] = value
                  end
                  if r == 46 && cc == 2
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["20"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["25"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["30"] = {}
                  end
                  if r == 46 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["30"][ltv_data] = value
                  end
                  if r == 47 && cc == 2
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["20"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["25"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["30"] = {}
                  end
                  if r == 47 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["30"][ltv_data] = value
                  end
                  if r == 48 && cc == 2
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["20"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["25"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["30"] = {}
                  end
                  if r == 48 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["30"][ltv_data] = value
                  end
                  if r == 50 && cc == 2
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["20"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["25"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["30"] = {}
                  end
                  if r == 50 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non Owner Occupied"]["30"][ltv_data] = value
                  end
                  if r >= 55 && r <= 60 && cc == 3
                    primary_key = get_value value
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"][primary_key] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 55 && r <= 60 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_arm_data[cc-2]
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"][primary_key][ltv_data] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"][primary_key][ltv_data] = value
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r == 61 && cc == 2
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"] = {}
                  end
                  if r == 61 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_arm_data[cc-2]
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = value
                  end
                  if r == 62 && cc == 2
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["15"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"] = {}
                  end
                  if r == 62 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_arm_data[cc-2]
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["15"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["15"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = value
                  end
                  if r == 44 && cc == 18
                    @other_hash["LoanType/Term"] = {}
                    @other_hash["LoanType/Term"]["Fixed"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["20"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["25"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["30"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["20"] = value
                    @other_hash["LoanType/Term"]["Fixed"]["25"] = value
                    @other_hash["LoanType/Term"]["Fixed"]["30"] = value
                  end
                  if r == 45 && cc == 18
                    @other_hash["LoanType/Term"]["Fixed"]["15"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["15"] = value
                  end
                  if r == 46 && cc == 18
                    @other_hash["LoanType/ArmBasic"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["5"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["5"] = value
                  end
                  if r == 47 && cc == 18
                    @other_hash["LoanType/ArmBasic"]["ARM"]["7"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["7"] = value
                  end
                  if r == 48 && cc == 18
                    @other_hash["LoanType/ArmBasic"]["ARM"]["10"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["10"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@jumbo_adjustment,@cash_out,@other_hash]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(@sheet)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def high_balance_extra
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "High Balance Extra")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @sub_hash = {}
        @cash_out = {}
        @bal_data = []
        @sub_data = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        key = ''
        bal_data = ''
        sub_data = ''
        @sheet = sheet
        (5..23).each do |r|
          row = sheet_data.row(r)
          if (row.compact.include?("High Balance Extra 30 Yr Fixed"))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              begin
                # title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                if @title.present?
                  # term
                  term = nil
                  if @title.include?("10yr") || @title.include?("10 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("15yr") || @title.include?("15 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("20yr") || @title.include?("20 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("25yr") || @title.include?("25 Yr")
                    term = @title.scan(/\d+/)[0]
                  elsif @title.include?("30yr") || @title.include?("30 Yr")
                    term = @title.scan(/\d+/)[0]
                  end


                  # rate type
                  if p_name.include?("Fixed")
                    loan_type = "Fixed"
                  elsif p_name.include?("ARM")
                    loan_type = "ARM"
                    arm_benchmark = "LIBOR"
                    arm_margin = 0
                  elsif p_name.include?("Floating")
                    loan_type = "Floating"
                  elsif p_name.include?("Variable")
                    loan_type = "Variable"
                  else
                    loan_type = "Fixed"
                  end
                  # rate arm
                  if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                    arm_basic = @title.scan(/\d+/)[0].to_i
                  end
                  # High Balance
                  if p_name.downcase.include?("high balance extra")
                    loan_size = "High-Balance Extra"
                  else
                    loan_size = "Conforming"
                  end
                  # loan_purpose
                  if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                    loan_purpose = "Refinance"
                  else
                    loan_purpose = "Purchase"
                  end
                  # lp and du
                  if p_name.downcase.include?('du ')
                    du = true
                  end
                  if p_name.downcase.include?('lp ')
                    lp = true
                  end
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program_ids << @program.id
                  @program.update(term: term,loan_type: loan_type, arm_basic: arm_basic, loan_category: @sheet_name, loan_size: loan_size, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                  @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (0..19).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
                        end
                        @data << value
                      end
                    end

                    if @data.compact.length == 0
                      break # terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                    @block_hash.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (25..44).each do |r|
          row = sheet_data.row(r)
          @bal_data = sheet_data.row(27)
          @sub_data = sheet_data.row(41)
          if row.compact.count >= 1
            (0..9).each do |max_column|
              cc = max_column
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Pricing Adjustments"
                    @adjustment_hash["LoanSize/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"] = {}
                  end
                  if value == "Cashout (adjustments are cumulative)"
                    @cash_out["RefinanceOption/FICO/LTV"] = {}
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                  end
                  if value == "Sub Financing (adjustments are cumulative)"
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  # All High Balance Extra Loans
                  if r >= 28 && r <= 32 && cc == 2
                    if value.include?(">")
                      ltv_key = value.tr('>=','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      ltv_key = get_value value
                    end
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"][ltv_key] = {}
                  end
                  if r >= 28 && r <= 32 && cc > 3 && cc <= 9
                    if @bal_data[cc-2].include?("<")
                      bal_data = "0-"+ @bal_data[cc-2].tr('<= ','')
                    else
                      bal_data = get_value @bal_data[cc-2]
                    end
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"][ltv_key][bal_data] = {}
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"][ltv_key][bal_data] = value
                  end
                  # Cashout Adjustments
                  if r >= 34 && r <= 38 && cc == 2
                    if value.include?(">")
                      ltv_key = value.tr('>=','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      ltv_key = get_value value
                    end
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key] = {}
                  end
                  if r >= 34 && r <= 38 && cc > 3 && cc <= 9
                    if @bal_data[cc-2].include?("<")
                      bal_data = "0-"+ @bal_data[cc-2].tr('<= ','')
                    else
                      bal_data = get_value @bal_data[cc-2]
                    end
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key][bal_data] = {}
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key][bal_data] = value
                  end

                  # Subordinate Financing Adjustments
                  if r >= 42 && r <= 44 && cc == 2
                    if value.include?("<")
                      ltv_key = "0-"+ value.tr('<= ','')
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key] = {}
                  end
                  if r >= 42 && r <= 44 && cc == 3
                    cltv_key = get_value value
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key] = {}
                  end
                  if r >= 42 && r <= 44 && cc > 3 && cc <= 5
                    if @sub_data[cc-2].include?(">")
                      sub_data = @sub_data[cc-2].tr('>= ','')+"-#{(Float::INFINITY).to_s.downcase}"
                    else
                      sub_data = get_value @sub_data[cc-2]
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_data] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_data] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@sub_hash,@cash_out]
        create_adjust(adjustment,sheet)
        create_program_association_with_adjustment(@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def freddie_arms
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Freddie ARMs")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Freddie Mac 10-1 ARM (5-2-5) Super Conforming"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                # title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                # term
                term = nil
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = 10
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = 15
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = 20
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = 25
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = 30
                end

                # rate type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # Arm Basic
                if @title.include?("3-1 ARM") || @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end

                # Arm Advanced
                if @title.downcase.include?("arm")
                  arm_advanced = @title.downcase.split("arm").last.tr('A-Za-z() ','')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end

                # conforming
                if p_name.downcase.include?("super conforming")
                  loan_size = "Super Conforming"
                elsif p_name.downcase.include?("conforming") 
                  conforming = true
                  loan_size = "Conforming"
                else
                  loan_size = "Conforming"
                end

                # freddie_mac
                freddie_mac = false
                if p_name.include?("Freddie Mac")
                  freddie_mac = true
                end

                # fannie_mae
                fannie_mae = false
                if p_name.include?("Fannie Mae") 
                  fannie_mae = true
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
                @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: @sheet_name, arm_advanced: arm_advanced,loan_size: loan_size, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (49..99).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(52)
          @lpmi = sheet_data.row(68)
          @fico = sheet_data.row(84)
          @unit = sheet_data.row(81)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/LTV/FICO"] = {}
                    @property_hash["LPMI/LTV/FICO"][true] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"] = {}
                    @property_hash["PropertyType"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"] = {}
                  end
                  if r >= 53 && r <= 59 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key] = {}
                  end
                  if r >= 53 && r <= 59 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = value
                  end
                  if r >= 63 && r <= 66 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 63 && r <= 66 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 69 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 69 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 70 && r <= 73 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 70 && r <= 73 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r == 74 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"] = {}
                  end
                  if r == 74 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = value
                  end
                  if r >= 76 && r <= 79 && cc == 5
                    primary_key = get_value value
                    @property_hash["LPMI/LTV/FICO"][true][primary_key] = {}
                  end
                  if r >= 76 && r <= 79 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 82 && r <= 83 && cc == 6
                    primary_key = value.split('s').first
                    @property_hash["PropertyType"][primary_key] = {}
                  end
                  if r >= 82 && r <= 83 && cc >= 9 && cc <= 11
                    unit_data = get_value @unit[cc-2]
                    @property_hash["PropertyType"][primary_key][unit_data] = {}
                    @property_hash["PropertyType"][primary_key][unit_data] = value
                  end
                  if r >= 85 && r <= 88 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 85 && r <= 88 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 85 && r <= 88 && cc >= 10 && cc <= 11
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 89 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-inf"] = value
                  end
                  if r == 90 && cc == 11
                    @property_hash["LTV"] = {}
                    @property_hash["LTV"]["90-Inf"] = {}
                    @property_hash["LTV"]["90-Inf"] = value
                  end
                  if r == 91 && cc == 11
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 92 && cc == 11
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r == 93 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 94 && r <= 96 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 82 && r <= 88 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 82 && r <= 88 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 82 && r <= 88 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 88 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 88 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 88 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 89 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 91 && cc == 19
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["0-75"] = {}
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["0-75"] = value
                  end
                  if r == 92 && cc == 19
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["75-Inf"] = {}
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["75-Inf"] = value
                  end
                  if r == 93 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption/LTV"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["0-75"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["0-75"] = value
                  end
                  if r == 93 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["75-Inf"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["75-Inf"] = value
                  end
                  if r == 96 && cc == 19
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def conforming_arms
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conforming ARMs")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || row.compact.include?("Fannie Mae 10-1 ARM (5-2-5) High Balance")
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                term = nil
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = @title.scan(/\d+/)[0]
                end

                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # rate arm
                if @title.include?("3-1 ARM") || @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end

                freddie_mac = false
                if p_name.downcase.include?("freddie mac")
                  freddie_mac = true
                end

                conforming = false
                if p_name.downcase.include?("conforming") 
                  conforming = true
                end

                fannie_mae = false
                if p_name.downcase.include?("fannie mae")
                  fannie_mae = true
                end
                # Arm Advanced
                if @title.downcase.include?("arm")
                  arm_advanced = @title.split("ARM").last.tr('A-Za-z ()', '')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end
                # High Balance
                if p_name.include?("High Balance")
                  loan_size = "High-Balance"
                else
                  loan_size = "Conforming"
                end

                # Fha, va, usda
                if p_name.downcase.include?("fha")
                  fha = true
                end
                if p_name.downcase.include?("va")
                  va = true
                end
                if p_name.downcase.include?("usda")
                  usda = true
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @program_ids << @program.id
     
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fha: fha, va: va, usda: usda, fannie_mae: fannie_mae, loan_size: loan_size, loan_category: @sheet_name, arm_basic: arm_basic, arm_advanced: arm_advanced ,base_rate: @block_hash, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)

                # @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (49..99).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(52)
          @lpmi = sheet_data.row(67)
          @fico = sheet_data.row(80)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/LTV/FICO"] = {}
                    @property_hash["LPMI/LTV/FICO"][true] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"] = {}
                    @property_hash["PropertyType"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Rate and Term"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Cash Out"] = {}
                  end
                  if r >= 53 && r <= 60 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key] = {}
                  end
                  if r >= 53 && r <= 60 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = value
                  end
                  if r >= 63 && r <= 65 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 63 && r <= 65 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 68 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 68 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 69 && r <= 72 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 69 && r <= 72 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 75 && r <= 78 && cc == 5
                    primary_key = get_value value
                    @property_hash["LPMI/LTV/FICO"][true][primary_key] = {}
                  end
                  if r >= 75 && r <= 78 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = value
                  end
                  
                  if r >= 81 && r <= 84 && cc == 5
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 81 && r <= 84 && cc == 6
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 81 && r <= 84 && cc >= 9 && cc <= 10
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r >= 87 && r <= 89 && cc == 6
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = {}
                    cc = cc + 3
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 91 && r <= 92 && cc == 5
                    primary_key = value.split('s').first
                    @property_hash["PropertyType"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType"][primary_key] = new_val
                  end
                  if r == 93 && cc == 5
                    @property_hash["PropertyType/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/LTV"]["Condo"]["75-Inf"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Condo"]["75-Inf"] = new_val
                  end
                  if r == 94 && cc == 9
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 95 && cc == 9
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r == 96 && cc == 9
                    @property_hash["LTV"] = {}
                    @property_hash["LTV"]["90-Inf"] = {}
                    @property_hash["LTV"]["90-Inf"] = value
                  end
                  if r >= 81 && r <= 87 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 81 && r <= 87 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 81 && r <= 87 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 87 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 87 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 87 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 88 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r >= 89 && r <= 90 && cc == 18
                    primary_key = get_value value
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Rate and Term"][primary_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Rate and Term"][primary_key] = new_val
                  end
                  if r >= 91 && r <= 92 && cc == 18
                    primary_key = get_value value
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Cash Out"][primary_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Cash Out"][primary_key] = new_val
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def homeready
    program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..76).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|

              cc = 3 + max_column*6 # (3 / 9 / 15) 3/8/13
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                # term
                term = nil
                program_heading = @title.split
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = @title.scan(/\d+/)[0]
                end

                # rate type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # rate arm
                if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end

                # Arm Advanced
                if @title.downcase.include?("arm") 
                  arm_advanced = @title.split("ARM").last.tr('A-Z- () ','')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end

                if p_name.downcase.include?("homeready")
                  fannie_mae_product = "HomeReady"
                end

                # fannie_mae
                if p_name.downcase.include?("fnma")
                  fannie_mae = true
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                loan_size = "Conforming"

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end

                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_ids << @program.id
                @program.update(term: term,loan_type: loan_type, arm_basic: arm_basic,arm_advanced: arm_advanced, fannie_mae_product: fannie_mae_product, loan_category: @sheet_name, loan_purpose: loan_purpose, du: du, lp: lp, loan_size: loan_size,fannie_mae: fannie_mae, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (73..131).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(80)
          @lpmi = sheet_data.row(97)
          @fico = sheet_data.row(112)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                  end
                  if value == "LPMI Adjustments Applied after Cap"
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  if value == "Adjustments Applied after Cap"
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 81 && r <= 88 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 81 && r <= 88 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 91 && r <= 93 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 91 && r <= 93 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 98 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 98 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 99 && r <= 100 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 99 && r <= 100 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 102 && r <= 105 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 102 && r <= 105 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 107 && r <= 110 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 107 && r <= 110 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 113 && r <= 118 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 113 && r <= 118 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 113 && r <= 118 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 119 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = value
                  end
                  if r == 120 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 121 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 122 && r <= 124 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"][primary_key] = new_val
                  end
                  if r == 126 && cc == 8
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r >= 114 && r <= 120 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 114 && r <= 120 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 114 && r <= 120 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 120 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 120 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 120 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 121 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 122 && cc == 19
                    @loan_amount["MiscAdjuster/LoanPurpose"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = value
                  end
                  if r == 123 && cc == 19
                    @loan_amount["PropertyType/LoanPurpose"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end
    create_program_association_with_adjustment(@sheet)
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def homeready_hb
    program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady HB")
        @sheet_name = sheet
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''

        (1..75).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                 # term
                term = nil
                if @title.include?("10yr") || @title.include?("10 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("15yr") || @title.include?("15 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("20yr") || @title.include?("20 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("25yr") || @title.include?("25 Yr")
                  term = @title.scan(/\d+/)[0]
                elsif @title.include?("30yr") || @title.include?("30 Yr")
                  term = @title.scan(/\d+/)[0]
                end

                # rate type
                if p_name.include?("Fixed")
                  loan_type = "Fixed"
                elsif p_name.include?("ARM")
                  loan_type = "ARM"
                  arm_benchmark = "LIBOR"
                  arm_margin = 0
                elsif p_name.include?("Floating")
                  loan_type = "Floating"
                elsif p_name.include?("Variable")
                  loan_type = "Variable"
                else
                  loan_type = "Fixed"
                end

                # Arm Basic
                if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM") || @title.include?("5/1 ARM") || @title.include?("7/1 ARM") || @title.include?("10/1 ARM")
                  arm_basic = @title.scan(/\d+/)[0].to_i
                end
                # Arm Advanced
                if @title.downcase.include?("arm") 
                  arm_advanced = @title.split("ARM").last.tr('A-Z- () ','')
                  if arm_advanced.include?('/')
                    arm_advanced = arm_advanced.tr('/','-')
                  else
                    arm_advanced
                  end
                end

                if p_name.include?("Fannie Mae") || p_name.downcase.include?("fnma")
                  fannie_mae = true
                end

                fannie_mae_home_ready = false
                if p_name.include?("Fannie Mae HomeReady")
                  fannie_mae_home_ready = true
                end
                # Loan Size
                if p_name.downcase.include?("high balance") || p_name.downcase.include?("hb")
                  loan_size = "High-Balance"
                else
                  loan_size = "Conforming"
                end
                # Fannie_mae_product
                if p_name.downcase.include?("homeready")
                  fannie_mae_product = 'HomeReady'
                end

                # loan_purpose
                if p_name.downcase.include?('refinance') || p_name.downcase.include?('refi')
                  loan_purpose = "Refinance"
                else
                  loan_purpose = "Purchase"
                end

                # lp and du
                if p_name.downcase.include?('du ')
                  du = true
                end
                if p_name.downcase.include?('lp ')
                  lp = true
                end
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_ids << @program.id
                @program.update(term: term,loan_type: loan_type, arm_basic: arm_basic, arm_advanced: arm_advanced, fannie_mae: fannie_mae, fannie_mae_home_ready: fannie_mae_home_ready, loan_category: @sheet_name,loan_size: loan_size, fannie_mae_product: fannie_mae_product, loan_purpose: loan_purpose, du: du, lp: lp, arm_benchmark: arm_benchmark, arm_margin: arm_margin)
                @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (77..130).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(79)
          @lpmi = sheet_data.row(96)
          @fico = sheet_data.row(111)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                  end
                  if value == "LPMI Adjustments Applied after Cap"
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @property_hash["LoanSize/LoanType/LTV"] = {}
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                  end
                  if value == "Adjustments Applied after Cap"
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 80 && r <= 87 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 80 && r <= 87 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 90 && r <= 92 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 90 && r <= 92 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 97 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 97 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 98 && r <= 99 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 98 && r <= 99 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 101 && r <= 104 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 101 && r <= 104 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 106 && r <= 109 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 106 && r <= 109 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 112 && r <= 117 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 112 && r <= 117 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 112 && r <= 117 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 118 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = value
                  end
                  if r == 119 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 120 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r == 121 && cc == 11
                    @property_hash["LoanSize/LoanType"] = {}
                    @property_hash["LoanSize/LoanType"]["High-Balance"] = {}
                    @property_hash["LoanSize/LoanType"]["High-Balance"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType"]["High-Balance"]["Fixed"] = value
                  end
                  if r >= 122 && r <= 123 && cc == 6
                    primary_key = get_value value
                    primary_key =  primary_key.gsub('--','-')
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][primary_key] = {}
                    cc = cc + 5
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][primary_key] = new_val
                  end
                  if r == 125 && cc == 8
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r >= 113 && r <= 119 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 113 && r <= 119 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 113 && r <= 119 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 119 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 119 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 119 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 120 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 121 && cc == 19
                    @loan_amount["MiscAdjuster/LoanPurpose"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = value
                  end
                  if r == 122 && cc == 19
                    @loan_amount["PropertyType/LoanPurpose"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        create_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private

  def read_sheet
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    @xlsx = Roo::Spreadsheet.open(file)
  end

  def check_sheet_empty
    action =  params[:action]
    sheet_data = @xlsx.sheet(action) rescue @xlsx.sheet(action.upcase) rescue @xlsx.sheet(action.downcase) rescue @xlsx.sheet(action.capitalize)

    if sheet_data.first_row.blank?
      @msg = "Sheet is empty."
      redirect_to ob_new_rez_wholesale5806_index_path
    end
  end

  def get_sheet
    @sheet_obj = Sheet.find(params[:id])
  end

  def get_program
    @program = Program.find(params[:id])
  end

  def get_bank
    @bank = Bank.find(params[:id])
  end

  def get_titles
    return ["FICO/LTV Adjustments - Loan Amount ≤ $1MM", "State Adjustments", "FICO/LTV Adjustments - Loan Amount > $1MM", "Feature Adjustments", "Max Price"]
  end

  def all_lp
    data = Adjustment::ALL_IP

    return data
  end

  def high_bal_adjustment
    data = Adjustment::HIGH_BALANCE_ADJUSTMENT
    return data
  end

  def jumbo_series_i_adjustment
      data = Adjustment::JUMBO_SERIES_I_ADJUSTMENT
    return data
  end

  def dream_big_adjustment
    data = Adjustment::DREAM_BIG_ADJUSTMENT

    return data
  end

  def table_data
    hash_keys = {
      "FICO/LTV Adjustments" => "LoanAmount/FICO/LTV",
      "Feature Adjustments"  => "Feature/LTV",
      "State Adjustments" => "State",
      "Max Price" => "Max Price"
    }

    return hash_keys
  end


  def make_adjust(block_hash, sheet)
    block_hash.keys.each do |key|
      unless ["Lender Paid MI Adj.", "Term/LTV/FICO"].include?(key)
        hash = {}
        hash[key] = block_hash[key]
        Adjustment.create(data: hash,loan_category: @sheet_name)
      else
        unless block_hash[key].empty?
          block_hash[key].keys.each do |s_key|
            h1 = {}
            h1[s_key] = block_hash[key][s_key]
            Adjustment.create(data: h1,loan_category: @sheet_name)
          end
        end
      end
    end
  end

  def find_key(title)
    if title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM") or title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
      base_key = table_data[@title.split(" -").first]
    else
      base_key = table_data[@title]
    end

    return base_key
  end

  def get_table_keys
    table_keys = Adjustment::MAIN_KEYS
    return table_keys
  end

  def get_value value1
    if value1.present?
      if value1.include?("<=") || value1.include?("<") || value1.include?("≤")
        value1 = "0-"+value1.split("<=").last.tr('A-Za-z%$><=≤, ','')
        value1 = value1.tr('–','-')
      elsif value1.include?(">") || value1.include?("+")
        value1 = value1.split(">").last.tr('A-Za-z+ ','')+"-Inf"
        value1 = value1.tr('–','-')
      elsif value1.include?("≥")
        value1 = value1.split("≥").last.tr('A-Za-z$, ','')+"-Inf"
        value1 = value1.tr('–','-')
      else
        value1 = value1.tr('$, ','')
        value1 = value1.tr('–','-')
      end
    end
  end

  def set_range value
    if value.split()[0].eql?("≤") || value.split()[0].eql?("<=") then
      value = "0-" + value.split()[1]
    elsif [">","≥",">=", "+"].include?(value.split()[0]) then
      value.split()[1] + "-#{Float::INFINITY}"
    elsif [">","≥",">=", "+"].include?(value.split("")[-1])
      value.split("+")[0] + "-#{Float::INFINITY}"
    elsif value.include?(">")
      value.split(">")[-1] + "-#{Float::INFINITY}"
    elsif value.include?("<=")
      value = "0-" + value.split("<=")[-1]
    end
  end

  def convert_range value
    if value.include?("≤") && value.include?("MM")
      value = "0-"+value.split("MM")[0].split("$")[-1] + "000000"
    elsif value.include?("-") && value.include?("MM") && value.include?("$")
      if value.split(" - ")[-1].split("MM")[0].split("$")[-1].include?(".")
        value = value.split(" - ")[0].split("MM")[0].split("$")[-1] + "000000" + "-" + value.split(" - ")[-1].split("MM")[0].split("$")[-1].gsub(".", "") + "000000"
      else
        value = value.split(" - ")[0].split("MM")[0].split("$")[-1].gsub(".", "") + "000000" + "-" + value.split(" - ")[-1].split("MM")[0].split("$")[-1] + "000000" if value.split(" - ")[0].split("MM")[0].split("$")[-1].include?(".")
      end
    end

    return value
  end

  def get_main_key heading
    heading.split(" ").each do |data|
      data.gsub!("#{data}", 'LoanType') if data.eql?("Fixed")
      data.gsub!("#{data}", 'Term') if data.eql?("terms")
    end
  end

  def create_program_association_with_adjustment(sheet)
    adjustment_list = Adjustment.where(loan_category: sheet)
    program_list = Program.where(loan_category: sheet)

    adjustment_list.each_with_index do |adj_ment, index|
      key_list = adj_ment.data.keys.first.split("/")
      program_filter1={}
      program_filter2={}
      include_in_input_values = false
      if key_list.present?
        key_list.each_with_index do |key_name, key_index|
          if (Program.column_names.include?(key_name.underscore))
            unless (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
              program_filter1[key_name.underscore] = nil
            else
              if (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
                program_filter2[key_name.underscore] = true
              end
            end
            include_in_input_values = true
          else
            if(Adjustment::INPUT_VALUES.include?(key_name))
              include_in_input_values = true
            end
          end
        end

        if (include_in_input_values)
          program_list1 = program_list.where.not(program_filter1)
          program_list2 = program_list1.where(program_filter2)

          if program_list2.present?
            program_list2.map{ |program| program.adjustments << adj_ment unless program.adjustments.include?(adj_ment) }
          end
        end
      end
    end
  end

  def create_adjust(block_hash, sheet)
    block_hash.each do |hash|
      if hash.present?
        hash.each do |key|
          data = {}
          data[key[0]] = key[1]
          Adjustment.create(data: data,loan_category: sheet)
        end
      end
    end
  end
end
