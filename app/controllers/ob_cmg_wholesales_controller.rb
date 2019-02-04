class ObCmgWholesalesController < ApplicationController
  before_action :get_sheet, only: [:gov, :agency, :durp, :oa, :jumbo_700,:jumbo_7200_6700, :jumbo_6600, :jumbo_6200, :jumbo_7600, :jumbo_6800, :jumbo_6900_7900, :programs, :jumbo_6400]
  before_action :get_program, only: [:single_program, :program_property]
  before_action :read_sheet, only: [:index,:gov, :agency, :durp, :oa, :jumbo_700,:jumbo_7200_6700, :jumbo_6600, :jumbo_6200, :jumbo_7600, :jumbo_6800, :jumbo_6900_7900, :programs, :jumbo_6400]

  def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "AGENCY")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "CMG Financial"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def gov
    @programs_ids = []
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "GOV")
          sheet_data = @xlsx.sheet(sheet)
          @programs_ids = []
          first_key = ''
          first_key1 = ''
          first_key2 = ''
          second_key = ''
          second_key1 = ''
          second_key2 = ''
          state_key = ''
          cc = ''
          ccc = ''
          cltv_key = ''
          k_val = ''
          key_val = ''
          value1 = ''
          @block_hash = {}
          @data_hash = {}
          @misc_hash = {}
          @state_hash = {}
          adj_key = []
          (10..60).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 4))
              rr = r + 1
              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 4*max_column + 1
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id              
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
          # Adjustments
          (67..87).each do |r|
            row = sheet_data.row(r)
            @key_data = sheet_data.row(40)
            if (row.compact.count >= 1)
              (0..7).each do |max_column|
                cc = max_column
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "GOVERNMENT ADJUSTMENTS"
                    first_key = "GovermentAdjustments"
                    @data_hash[first_key] = {}
                  end

                  if value == "FICO, LOAN AMOUNT & PROPERTY TYPE ADJUSTMENTS"
                    second_key = "FicoLoanAmont"
                    @data_hash[first_key][second_key] = {}
                  end

                  if r >= 70 && r <= 87 && cc == 1
                    value = get_value value
                    ccc = cc + 6
                    c_val = sheet_data.cell(r,ccc)
                    @data_hash[first_key][second_key][value] = c_val
                  end

                end
              end

              (10..16).each do |max_column|
                cc = max_column
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "MISCELLANEOUS"
                    first_key1 = "GovermentAdjustments"
                    second_key1 = "Miscellaneous"
                    @misc_hash[first_key1] = {}
                    @misc_hash[first_key1][second_key1] = {}
                  end

                  if r >= 70 && r <= 77 && cc == 10
                    value1 = get_key value
                    ccc = cc + 6
                    c_val = sheet_data.cell(r,ccc)
                    @misc_hash[first_key1][second_key1][value] = c_val
                  end

                  if value == "STATE ADJUSTMENTS"
                    first_key2 = "GovermentAdjustments"
                    second_key2 = "StateAdjustments"
                    @state_hash[first_key2] = {}
                    @state_hash[first_key2][second_key2] = {}
                  end

                  if r >= 80 && r <= 87 && cc == 11
                    adj_key = value.split(', ')
                    adj_key.each do |f_key|
                      key_val = f_key
                      ccc = cc + 5
                      k_val = sheet_data.cell(r,ccc)
                      @state_hash[first_key2][second_key2][key_val] = k_val
                    end
                  end

                end
              end
            end
          end
          adjustment = [@data_hash,@misc_hash,@state_hash]
          make_adjust(adjustment,sheet)
          create_program_association_with_adjustment(sheet)
        end
<<<<<<< HEAD
=======
        adjustment = [@data_hash,@misc_hash,@state_hash]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
>>>>>>> 69f130303a58615f3be124e9326881b9f2307aa4
      end
    rescue
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def agency
    @programs_ids = []
    begin
    @xlsx.sheets.each do |sheet|
      # Programs
      if (sheet == "AGENCY")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        (10..87).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              if @title.present? && @title != "2.250% MARGIN - 2/2/6 CAPS - 1 YR LIBOR" && @title != "2.250% MARGIN - 5/2/5 CAPS - 1 YR LIBOR"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
      # Adjustments
      if (sheet == "AGENCYLLPAS")
        sheet_data = @xlsx.sheet(sheet)
        @ltv_data = []
        @cltv_data = []
        @adjustment_hash = {}
        @cashout_adjustment = {}
        @subordinate_hash = {}
        @adjustment_cap = {}
        @loan_adjustment = {}
        @state_adjustments = {}
        @other_adjustment = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key = ''
        secondary_key1 = ''
        ltv_key = ''
        cltv_key = ''
        cash_key = ''
        key = ''
        (8..62).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(10)
          @cltv_data = sheet_data.row(38)
          (0..16).each do |cc|
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == "AGENCY FIXED AND ARM ADJUSTMENTS"
                primary_key = "LoanType/Term/FICO/LTV"
                @adjustment_hash[primary_key] = {}
                cash_key = "CashOut/FICO/LTV"
                @cashout_adjustment[cash_key] = {}
              end
              if value == "SUBORDINATE FINANCING"
                primary_key = "FinancingType/LTV/CLTV/FICO"
                @subordinate_hash[primary_key] = {}
              end
              if value == "HOMEREADY ADJUSTMENT CAPS*"
                primary_key = "FinancingType/LoanType"
                @adjustment_cap[primary_key] = {}
              end
              # AGENCY FIXED AND ARM ADJUSTMENTS
              if r >= 11 && r <= 24  && cc == 1
                if value.include?("Condo")
                  secondary_key = "LoanType/Fixed/Condo"
                else
                  secondary_key = get_value value
                end
                @adjustment_hash[primary_key][secondary_key] = {}
              end
              if r >= 11 && r <= 24 && cc >= 9 && cc <= 16
                ltv_key = get_value @ltv_data[cc-1]
                @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                @adjustment_hash[primary_key][secondary_key][ltv_key] = value
              end

              # cashout_adjustment
              if r >= 25 && r <= 31 && cc == 1
                secondary_key = get_value value
                @cashout_adjustment[cash_key][secondary_key] = {}
              end
              if r >= 25 && r <= 31 && cc >= 9 && cc <= 16
                ltv_key = @ltv_data[cc-1]
                @cashout_adjustment[cash_key][secondary_key][ltv_key] = {}
                @cashout_adjustment[cash_key][secondary_key][ltv_key] = value
              end

              # High Balance Adjustments
              if r >= 33 && r <= 35 && cc == 1
                unless @adjustment_hash[primary_key].has_key?("ARM")
                  secondary_key = value.split.first
                else
                  secondary_key = "/" + value.split.first
                end
                @adjustment_hash[primary_key][secondary_key] = {}
              end
              if r >= 33 && r <= 35 && cc >= 9 && cc <= 16
                ltv_key = @ltv_data[cc-1]
                @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                @adjustment_hash[primary_key][secondary_key][ltv_key] = value
              end

              # SUBORDINATE FINANCING
              if r >= 39 && r <= 43 && cc == 1
                secondary_key = get_value value
                @subordinate_hash[primary_key][secondary_key] = {}
              end
              if r >= 39 && r <= 43 && cc == 3
                ltv_key = get_value value
                @subordinate_hash[primary_key][secondary_key][ltv_key] = {}
              end
              if r >= 39 && r <= 43 && cc >= 5 && cc <= 7
                cltv_key = get_value @cltv_data[cc-1]
                @subordinate_hash[primary_key][secondary_key][ltv_key][cltv_key] = {}
                @subordinate_hash[primary_key][secondary_key][ltv_key][cltv_key] = value
              end
              if r == 44 && cc == 1
                secondary_key = "Home Possible"
                @subordinate_hash[primary_key][secondary_key] = {}
              end
              if r == 44 && cc == 5
                @subordinate_hash[primary_key][secondary_key] = value
              end

              # HOMEREADY ADJUSTMENT CAPS
              if r >= 48 && r <= 49 && cc ==1
                if value == "LTV > 80% AND FICO >= 680"
                  secondary_key = "FICO/LTV"
                else
                  secondary_key = "All/FICO/LTV"
                end
                @adjustment_cap[primary_key][secondary_key] = {}
              end
              if r >= 48 && r <= 49 && cc == 8
                @adjustment_cap[primary_key][secondary_key] = value
              end
              if r >= 54 && r <= 55 && cc ==1
                if value == "LTV > 80% AND FICO >= 680"
                  secondary_key = "HOME/FICO/LTV"
                else
                  secondary_key = "All/HOME/FICO/LTV"
                end
                @adjustment_cap[primary_key][secondary_key] = {}
              end
              if r >= 54 && r <= 55 && cc == 8
                @adjustment_cap[primary_key][secondary_key] = value
              end
            end
          end
          (10..16).each do |cc|
            value = sheet_data.cell(r,cc)
            if value.present?
              if value == "LOAN AMOUNT "
                primary_key1 = "LoanType/LoanAmount/CLTV"
                @loan_adjustment[primary_key1] = {}
              end
              if value == "STATE ADJUSTMENTS"
                primary_key1 = "State"
                @state_adjustments[primary_key1] = {}
              end
              if value == "MISCELLANEOUS"
                primary_key1 = "Miscellaneous"
                @other_adjustment[primary_key1] = {}
              end
              # LOAN AMOUNT
              if r >= 38 && r <= 42 && cc == 10
                secondary_key1 = get_value value
                @loan_adjustment[primary_key1][secondary_key1] = {}
              end
              if r >= 38 && r <= 42 && cc == 16
                @loan_adjustment[primary_key1][secondary_key1] = value
              end
              # STATE ADJUSTMENTS
              if r >= 46 && r <= 52 && cc == 11
                adj_key = value.split(', ')
                adj_key.each do |f_key|
                  key = f_key
                  ccc = cc + 5
                  c_val = sheet_data.cell(r,ccc)
                  @state_adjustments[primary_key1][key] = c_val
                end
              end
              # MISCELLANEOUS
              if r >= 55 && r <= 62 && cc == 10
                secondary_key1 = value
                @other_adjustment[primary_key1][secondary_key1] = {}
              end
              if r >= 55 && r <= 62 && cc == 16
                @other_adjustment[primary_key1][secondary_key1] = value
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@cashout_adjustment,@subordinate_hash,@adjustment_cap,@loan_adjustment,@state_adjustments,@other_adjustment]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
      end
    end
    rescue
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def durp
    @programs_ids = []
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "DURP")
          sheet_data = @xlsx.sheet(sheet)
          @programs_ids = []
          @adjustment_hash = {}
          @subordinate_hash = {}
          @adjustment_cap = {}
          @misc_adjustment = {}
          @state_adjustment = {}
          primary_key = ''
          primary_key1 = ''
          secondary_key = ''
          fnma_key = ''
          sub_data = ''
          sub_key = ''
          cltv_key = ''
          cap_key = ''
          m_key = ''
          key = ''
          loan_key = ''
          (10..53).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 4))
              rr = r + 1
              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 4*max_column + 1

                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
          # Adjustment
          (55..88).each do |r|
            row = sheet_data.row(r)
            @fnma_data = sheet_data.row(57)
            @sub_data = sheet_data.row(73)
            @cap_data = sheet_data.row(80)
            if row.compact.count >= 1
              (0..16).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?

                  if value == "FNMA DU REFI PLUS ADJUSTMENTS"
                    primary_key = "FNMA/DU"
                    @adjustment_hash[primary_key] = {}
                  elsif value == "SUBORDINATE FINANCING"
                    primary_key = "FinancingType/FICO/LTV/CLTV"
                    @subordinate_hash[primary_key] = {}
                  elsif value == "DU REFI PLUS ADJUSTMENT CAP (MAX ADJ) *"
                    primary_key = value
                    @adjustment_cap[primary_key] = {}
                  end
                  if r >= 58 && r <= 70 && cc == 1
                    secondary_key = get_value value
                    @adjustment_hash[primary_key][secondary_key] = {}
                  end
                  if r >= 58 && r <= 70 && cc >= 8 && cc <= 16
                    fnma_key = get_value @fnma_data[cc-1]
                    @adjustment_hash[primary_key][secondary_key][fnma_key] = {}
                    @adjustment_hash[primary_key][secondary_key][fnma_key] = value
                  end

                  # subordinate adjustment
                  if r >= 74 && r <= 78 && cc == 1
                    secondary_key = get_value value
                    @subordinate_hash[primary_key][secondary_key] = {}
                  end
                  if r >= 74 && r <= 78 && cc == 3
                    cltv_key = get_value value
                    @subordinate_hash[primary_key][secondary_key][cltv_key] = {}
                  end
                  if r >= 74 && r <= 78 && cc >= 5 && cc <= 7
                    sub_data = get_value @sub_data[cc-1]
                    @subordinate_hash[primary_key][secondary_key][cltv_key][sub_data] = {}
                    @subordinate_hash[primary_key][secondary_key][cltv_key][sub_data] = value
                  end
                  # Adjustment Cap
                  if r >= 81 && r <= 83 && cc == 1
                    secondary_key = value
                    @adjustment_cap[primary_key][secondary_key] = {}
                  end
                  if r >= 81 && r <= 83 && cc == 4
                    cltv_key = get_value value
                    @adjustment_cap[primary_key][secondary_key][cltv_key] = {}
                  end
                  if r >= 81 && r <= 83 && cc >= 5 && cc <= 7
                    cap_key = get_value @cap_data[cc-1]
                    @adjustment_cap[primary_key][secondary_key][cltv_key][cap_key] = {}
                    @adjustment_cap[primary_key][secondary_key][cltv_key][cap_key] = value
                  end
                end
              end
              (10..16).each do |cc|
                value = sheet_data.cell(r,cc)
                if value == "MISCELLANEOUS"
                  primary_key1 = "Miscellaneous"
                  @misc_adjustment[primary_key1] = {}
                end
                if value == "LOAN AMOUNT "
                  primary_key1 = "LoanType/LoanAmount/CLTV"
                  @misc_adjustment[primary_key1] = {}
                end
                if value == "STATE ADJUSTMENTS"
                  primary_key1 = "State"
                  @state_adjustment[primary_key1] = {}
                end
                if value.present?
                  # MISCELLANEOUS
                  if r >= 73 && r <= 74 && cc == 10
                    m_key = value
                    @misc_adjustment[primary_key1][m_key] = {}
                  end
                  if r >= 73 && r <= 74 && cc == 16
                    @misc_adjustment[primary_key1][m_key] = value
                  end
                  # LOAN AMOUNT ADJUSTMENT
                  if r >= 76 && r <= 80 && cc == 10
                    # m_key = value
                    m_key =  value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                    @misc_adjustment[primary_key1][m_key] = {}
                  end
                  if r >= 76 && r <= 80 && cc == 16
                    @misc_adjustment[primary_key1][m_key] = value
                  end
                  # STATE ADJUSTMENTS
                  if r >= 83 && r <= 88 && cc == 11
                    adj_key = value.split(', ')
                    adj_key.each do |f_key|
                      key = f_key
                      ccc = cc + 5
                      c_val = sheet_data.cell(r,ccc)
                      @state_adjustment[primary_key1][key] = c_val
                    end
                  end
                end
              end
            end
          end
          adjustment = [@adjustment_hash,@subordinate_hash,@adjustment_cap,@misc_adjustment,@state_adjustment]
          make_adjust(adjustment,sheet)

          create_program_association_with_adjustment(sheet)
        end
      end
    rescue
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def oa
    @programs_ids = []
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "OA")
          sheet_data = @xlsx.sheet(sheet)
          @programs_ids = []
          @adjustment_hash = {}
          @subordinate_hash = {}
          @adjustment_cap = {}
          @misc_adjustment = {}
          @state_adjustment = {}
          primary_key = ''
          primary_key1 = ''
          secondary_key = ''
          fnma_key = ''
          sub_data = ''
          sub_key = ''
          cltv_key = ''
          cap_key = ''
          m_key = ''
          key = ''
          loan_key = ''
          @programs_ids = []
          (10..53).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 4))
              rr = r + 1
              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 4*max_column + 1

                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
          # Adjustment
          (54..88).each do |r|
            row = sheet_data.row(r)
            @fnma_data = sheet_data.row(56)
            @sub_data = sheet_data.row(73)
            @cap_data = sheet_data.row(82)
            if row.compact.count >= 1
              (0..16).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?

                  if value == "FHLMC LP OPEN ACCESS ADJUSTMENTS"
                    primary_key = "FHLMC/LP"
                    @adjustment_hash[primary_key] = {}
                  elsif value == "SUBORDINATE FINANCING"
                    primary_key = "FinancingType/FICO/LTV/CLTV"
                    @subordinate_hash[primary_key] = {}
                  elsif value == "OPEN ACCESS ADJUSTMENT CAP (MAX ADJ) *"
                    primary_key = value
                    @adjustment_cap[primary_key] = {}
                  end
                  if r >= 57 && r <= 70 && cc == 1
                    secondary_key = get_value value
                    @adjustment_hash[primary_key][secondary_key] = {}
                  end
                  if r >= 57 && r <= 70 && cc >= 8 && cc <= 16
                    fnma_key = get_value @fnma_data[cc-1]
                    @adjustment_hash[primary_key][secondary_key][fnma_key] = {}
                    @adjustment_hash[primary_key][secondary_key][fnma_key] = value
                  end

                  # subordinate adjustment
                  if r >= 74 && r <= 80 && cc == 1
                    secondary_key = get_value value
                    @subordinate_hash[primary_key][secondary_key] = {}
                  end
                  if r >= 74 && r <= 80 && cc == 3
                    cltv_key = get_value value
                    @subordinate_hash[primary_key][secondary_key][cltv_key] = {}
                  end
                  if r >= 74 && r <= 80 && cc >= 5 && cc <= 7
                    sub_data = get_value @sub_data[cc-1]
                    @subordinate_hash[primary_key][secondary_key][cltv_key][sub_data] = {}
                    @subordinate_hash[primary_key][secondary_key][cltv_key][sub_data] = value
                  end
                  # Adjustment Cap
                  if r >= 83 && r <= 85 && cc == 1
                    secondary_key = value
                    @adjustment_cap[primary_key][secondary_key] = {}
                  end
                  if r >= 83 && r <= 85 && cc == 4
                    cltv_key = get_value value
                    @adjustment_cap[primary_key][secondary_key][cltv_key] = {}
                  end
                  if r >= 83 && r <= 85 && cc >= 5 && cc <= 7
                    cap_key = get_value @cap_data[cc-1]
                    @adjustment_cap[primary_key][secondary_key][cltv_key][cap_key] = {}
                    @adjustment_cap[primary_key][secondary_key][cltv_key][cap_key] = value
                  end
                end
              end
              (10..16).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "MISCELLANEOUS"
                    primary_key1 = "Miscellaneous"
                    @misc_adjustment[primary_key1] = {}
                  end
                  if value == "LOAN AMOUNT "
                    primary_key1 = "LoanType/LoanAmount/CLTV"
                    @misc_adjustment[primary_key1] = {}
                  end
                  if value == "STATE ADJUSTMENTS"
                    primary_key1 = "State"
                    @state_adjustment[primary_key1] = {}
                  end

                  # MISCELLANEOUS
                  if r >= 73 && r <= 74 && cc == 10
                    m_key = value
                    @misc_adjustment[primary_key1][m_key] = {}
                  end
                  if r >= 73 && r <= 74 && cc == 16
                    @misc_adjustment[primary_key1][m_key] = value
                  end
                  # LOAN AMOUNT ADJUSTMENT
                  if r >= 76 && r <= 80 && cc == 10
                    m_key =  value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                    @misc_adjustment[primary_key1][m_key] = {}
                  end
                  if r >= 76 && r <= 80 && cc == 16
                    @misc_adjustment[primary_key1][m_key] = value
                  end
                  # STATE ADJUSTMENTS
                  if r >= 83 && r <= 88 && cc == 11
                    adj_key = value.split(', ')
                    adj_key.each do |f_key|
                      key = f_key
                      ccc = cc + 5
                      c_val = sheet_data.cell(r,ccc)
                      @state_adjustment[primary_key1][key] = c_val
                    end
                  end
                end
              end
            end
          end
          adjustment = [@adjustment_hash,@subordinate_hash,@adjustment_cap,@misc_adjustment,@state_adjustment]
          make_adjust(adjustment,sheet)

          create_program_association_with_adjustment(sheet)
        end
      end
    rescue
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def jumbo_700
    @programs_ids = []
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "JUMBO 700")
          sheet_data = @xlsx.sheet(sheet)
          @programs_ids = []
          first_key = ''
          @adjustment_hash = {}
          @state_adjustment = {}
          @data_hash = {}
          @sheet = sheet
          @ltv_data = []
          key = ''
          key1 = ''
          key2 = ''
          key3 = ''
          ltv_key = ''
          cltv_key = ''
          c_val = ''
          cc = ''
          value = ''
          state_key = ''
          adj_key = []
          (10..21).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 4))
              rr = r + 1
              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 4*max_column + 1

                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end

          # adjustments
          (23..47).each do |r|
            row = sheet_data.row(r)
            @ltv_data = sheet_data.row(25)
            if (row.compact.count > 1)
              (0..9).each do |max_column|
                cc = max_column
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "ELITE JUMBO 700 SERIES ADJUSTMENTS"
                    first_key = "LoanAmount/FICO/LTV"
                    key = first_key
                    @adjustment_hash[key] = {}
                  end
                  if value == "Loan Amount <= $1,000,000"
                    key1 = "0"
                    @adjustment_hash[key][key1] = {}
                  end
                  if value == "Loan Amount > $1,000,000"
                    key1 = "$1,000,000"
                    @adjustment_hash[key][key1] = {}
                  end
                  if r >= 27 && r <= 33 && cc == 1
                    ltv_key = get_value value
                    @adjustment_hash[key][key1][ltv_key] = {}
                  end
                  if r >= 27 && r <= 33 && cc > 4 && cc <= 9
                    cltv_key = get_value @ltv_data[cc-1]
                    @adjustment_hash[key][key1][ltv_key][cltv_key] = value
                  end
                  if r >= 35 && r <= 47 && cc == 1
                    ltv_key = get_value value
                    @adjustment_hash[key][key1][ltv_key] = {}
                  end
                  if r >= 35 && r <= 47 && cc >= 4 && cc <= 9
                    cltv_key = get_value @ltv_data[cc-1]
                    @adjustment_hash[key][key1][ltv_key][cltv_key] = value
                  end
                end
              end

              #For STATE ADJUSTMENTS
              (12..16).each do |max_column|
                cc = max_column
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "STATE ADJUSTMENTS"
                    state_key = "State"
                    @state_adjustment[state_key] = {}
                  end
                  if r >= 24 && r < 28 && cc == 12
                    adj_key = value.split(', ')
                    adj_key.each do |f_key|
                      key3 = f_key
                      ccc = cc + 4
                      c_val = sheet_data.cell(r,ccc)
                      @state_adjustment[state_key][key3] = c_val
                    end
                  end
                end
              end

              #For MISCELLANEOUS
              (12..16).each do |max_column|
                cc = max_column
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "MISCELLANEOUS"
                    state_key = "Miscellaneous"
                    @adjustment_hash[state_key] = {}
                  end
                  if r == 31 && cc == 12
                    ccc = cc + 4
                    c_val = sheet_data.cell(r,ccc)
                    @adjustment_hash[state_key][value] = c_val
                  end
                end
              end
            end
          end
          adjustment = [@adjustment_hash,@state_adjustment]
          make_adjust(adjustment,sheet)

          create_program_association_with_adjustment(sheet)
        end
      end
    rescue
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def jumbo_6200
    @programs_ids = []
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "JUMBO 6200")
          sheet_data = @xlsx.sheet(sheet)
          first_key = ''
          second_key = ''
          cc = ''
          cltv_key = ''
          key_val = ''
          @data_hash = {}
          @key_data = []
          @key2_data = []
          @programs_ids = []
          (10..34).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 4))
              rr = r + 1
              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 4*max_column + 1

                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
          # Adjustments
          (37..88).each do |r|
            row = sheet_data.row(r)
            @key_data = sheet_data.row(40)
            @key2_data = sheet_data.row(83)
            if (row.compact.count >= 1)
              (0..13).each do |max_column|
                cc = max_column
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "PREMIER JUMBO 6200 SERIES ADJUSTMENTS"
                    first_key = "LoanPurpose/FICO/LTV"
                    @data_hash[first_key] = {}
                  end
                  if value == "Purchase Transaction"
                    second_key = "Purchase"
                    @data_hash[first_key][second_key] = {}
                  end
                  if value == "Rate/Term Transaction"
                    second_key = "Rate/Term"
                    @data_hash[first_key][second_key] = {}
                  end
                  if value == "Cash Out Transaction"
                    second_key = "CashOut"
                    @data_hash[first_key][second_key] = {}
                  end

                  # Purchase Transaction Adjustment
                  if r >= 41 && r <= 46 && cc == 1
                    cltv_key = get_value value
                    @data_hash[first_key][second_key][cltv_key] = {}
                  end
                  if r >= 41 && r <= 46 && cc >= 6 && cc <= 13
                    key_val = get_value @key_data[cc-1]
                    @data_hash[first_key][second_key][cltv_key][key_val] = value
                  end

                  # Rate/Term Transaction Adjustment
                  if r >= 49 && r <= 54 && cc == 1
                    cltv_key = get_value value
                    @data_hash[first_key][second_key][cltv_key] = {}
                  end
                  if r >= 49 && r <= 54 && cc >= 6 && cc <= 13
                    key_val = get_value @key_data[cc-1]
                    @data_hash[first_key][second_key][cltv_key][key_val] = value
                  end

                  # Cash Out Transaction Adjustment
                  if r >= 57 && r <= 77 && cc == 1
                    cltv_key = get_value value
                    @data_hash[first_key][second_key][cltv_key] = {}
                  end
                  if r >= 57 && r <= 77 && cc >= 6 && cc <= 13
                    key_val = get_value @key_data[cc-1]
                    @data_hash[first_key][second_key][cltv_key][key_val] = value
                  end

                  # MISCELLANEOUS Adjustment
                  if value == "MISCELLANEOUS"
                    second_key = "Miscellaneous"
                    @data_hash[first_key][second_key] = {}
                    k_val = sheet_data.cell(r+1,cc)
                    v_val = sheet_data.cell(r+1,cc+3)
                    @data_hash[first_key][second_key][k_val] = v_val
                  end

                  # MAX PRICE AFTER ADJUSTMENTS
                  if value == "MAX PRICE AFTER ADJUSTMENTS"
                    second_key = "MaxPriceAfterAdjustments"
                    @data_hash[first_key][second_key] = {}
                  end
                  if r >= 84 && r <= 85 && cc == 1
                    cltv_key = value
                    @data_hash[first_key][second_key][cltv_key] = {}
                  end
                  if r >= 84 && r <= 85 && cc >= 2 && cc <= 4
                    key_val = @key2_data[cc-1]
                    @data_hash[first_key][second_key][cltv_key][key_val] = value
                  end
                end
              end
            end
          end
          adjustment = [@data_hash]
          make_adjust(adjustment,sheet)

          create_program_association_with_adjustment(sheet)
        end
      end
    rescue
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def jumbo_7200_6700
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 7200 & 6700")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @purchase_adjustment = {}
        @rate_adjustment = {}
        @other_adjustment = {}
        @jumbo_purchase_adjustment = {}
        @jumbo_rate_adjustment = {}
        @jumbo_other_adjustment = {}
        @cltv_data = []
        @ltv_data = []
        primary_key = ''
        secondary_key = ''
        cltv_key = ''
        m_key = ''
        max_key = ''
        ltv_key = ''
        key = ''
        adj_key = ''
        (10..22).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              if cc < 5
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                  (1..50).each do |max_row|
                    @data = []
                    (0..3).each_with_index do |index, c_i|
                      rrr = rr + max_row -1
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        elsif (c_i == 1)
                          @block_hash[key][21] = value
                        elsif (c_i == 2)
                          @block_hash[key][30] = value
                        elsif (c_i == 3)
                          @block_hash[key][45] = value
                        end
                        @data << value
                      end
                    end
                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                end
                if  @block_hash.values.first.values.first == "21 Day"
                  if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
          (56..64).each do |r|
            row = sheet_data.row(r)
            if ((row.compact.count > 1) && (row.compact.count <= 4))
              rr = r + 1
              max_column_section = row.compact.count - 1
              (0..max_column_section).each do |max_column|
                cc = 4*max_column + 1

                @title = sheet_data.cell(r,cc)
                if cc < 5
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  program_property
                  @programs_ids << @program.id

                  # @program.adjustments.destroy_all
                  @block_hash = {}
                  key = ''
                  (1..50).each do |max_row|
                    @data = []
                    (0..3).each_with_index do |index, c_i|
                      rrr = rr + max_row -1
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        elsif (c_i == 1)
                          @block_hash[key][21] = value
                        elsif (c_i == 2)
                          @block_hash[key][30] = value
                        elsif (c_i == 3)
                          @block_hash[key][45] = value
                        end
                        @data << value
                      end
                    end
                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                end
                if  @block_hash.values.first.values.first == "21 Day"
                  if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
          (12..46).each do |r|
            row = sheet_data.row(r)
            @cltv_data = sheet_data.row(13)
            if row.compact.count >= 1
              (6..16).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Purchase Transaction"
                    primary_key = "LoanPurpose/FICO/LTV"
                    @purchase_adjustment[primary_key] = {}
                  elsif value == "Rate/Term Transaction"
                    primary_key = "LoanType/Term/FICO/LTV"
                    @rate_adjustment[primary_key] = {}
                  elsif value == "Cash Out Transaction"
                    primary_key = "LoanPurpose/RefinanceOption/LTV"
                    @adjustment_hash[primary_key] = {}
                  end
                  # Purchase Transaction Adjustment
                  if r >= 14 && r <= 19 && cc == 6
                    secondary_key = get_value value
                    @purchase_adjustment[primary_key][secondary_key] = {}
                  end
                  if r >= 14 && r <= 19 && cc >= 10 && cc <= 16
                    cltv_key = get_value @cltv_data[cc-1]
                    @purchase_adjustment[primary_key][secondary_key][cltv_key] = {}
                    @purchase_adjustment[primary_key][secondary_key][cltv_key] = value
                  end

                  # Rate/Term Transaction Adjustment
                  if r >= 22 && r <= 27 && cc == 6
                    secondary_key = get_value value
                    @rate_adjustment[primary_key][secondary_key] = {}
                  end
                  if r >= 22 && r <= 27 && cc >= 10 && cc <= 16
                    cltv_key = get_value @cltv_data[cc-1]
                    @rate_adjustment[primary_key][secondary_key][cltv_key] = {}
                    @rate_adjustment[primary_key][secondary_key][cltv_key] = value
                  end

                  # Cash Out Transaction Adjustment
                  if r >= 30 && r <= 46 && cc == 6
                    if value.include?("Loan Amount")
                      secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                    else
                      secondary_key = get_value value
                    end
                    @adjustment_hash[primary_key][secondary_key] = {}
                  end
                  if r >= 30 && r <= 48 && cc >= 10 && cc <= 16
                    cltv_key = get_value @cltv_data[cc-1]
                    @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                    @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                  end
                end
              end
              (1..4).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "MAX PRICE AFTER ADJUSTMENTS"
                    max_key = "LoanType/LA/"
                    @other_adjustment[max_key] = {}
                  end
                  # MISCELLANEOUS
                  if r == 25 && cc == 1
                    m_key = "Miscellaneous/NY"
                    @other_adjustment[m_key] = {}
                  end
                  if r == 25 && cc == 4
                    @other_adjustment[m_key] = value
                  end
                  # MAX PRICE AFTER ADJUSTMENTS
                  if r >= 29 && r <= 30 && cc == 1
                    key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
                    @other_adjustment[max_key][key] = {}
                  end
                  if r >= 29 && r <= 30 && cc == 4
                    @other_adjustment[max_key][key] = value
                  end
                end
              end
            end
          end
          adjustment = [@purchase_adjustment,@rate_adjustment,@adjustment_hash,@other_adjustment]
          make_adjust(adjustment,sheet)
          (56..77).each do |r|
            row = sheet_data.row(r)
            @ltv_data = sheet_data.row(59)
            if row.compact.count >= 1
              (10..16).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Purchase Transaction"
                    primary_key = "Jumbo/LoanPurpose/FICO/LTV"
                    @jumbo_purchase_adjustment[primary_key] = {}
                  elsif value == "Rate/Term Transaction"
                    primary_key = "Jumbo/LoanType/Term/FICO/LTV"
                    @jumbo_rate_adjustment[primary_key] = {}
                  elsif value == "MISCELLANEOUS"
                    primary_key = "Jumbo/NY/FICO/LTV"
                    @jumbo_other_adjustment[primary_key] = {}
                  end
                  # Purchase Transaction Adjustment
                  if r >= 60 && r <= 63 && cc == 10
                    secondary_key = get_value value
                    @jumbo_purchase_adjustment[primary_key][secondary_key] = {}
                  end
                  if r >= 60 && r <= 63 && cc >= 15 && cc <= 16
                    ltv_key = get_value @ltv_data[cc-1]
                    @jumbo_purchase_adjustment[primary_key][secondary_key][ltv_key] = {}
                    @jumbo_purchase_adjustment[primary_key][secondary_key][ltv_key] = value
                  end

                  # Rate/Term Transaction Adjustment
                  if r >= 66 && r <= 71 && cc == 10
                    if value.include?("Loan Amount")
                      secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                    else
                      secondary_key = get_value value
                    end
                    @jumbo_rate_adjustment[primary_key][secondary_key] = {}
                  end
                  if r >= 66 && r <= 71 && cc >= 15 && cc <= 16
                    ltv_key = get_value @ltv_data[cc-1]
                    @jumbo_rate_adjustment[primary_key][secondary_key][ltv_key] = {}
                    @jumbo_rate_adjustment[primary_key][secondary_key][ltv_key] = value
                  end

                  if r == 72 && cc == 10
                    key = "FL"
                    adj_key = "NV"
                    @jumbo_rate_adjustment[primary_key][key] = {}
                    @jumbo_rate_adjustment[primary_key][adj_key] = {}
                  end
                  if r == 72 && cc >= 15 && cc <= 16
                    @jumbo_rate_adjustment[primary_key][key][@ltv_data[cc-1]] = {}
                    @jumbo_rate_adjustment[primary_key][adj_key][@ltv_data[cc-1]] = {}
                    @jumbo_rate_adjustment[primary_key][key][@ltv_data[cc-1]] = value
                    @jumbo_rate_adjustment[primary_key][adj_key][@ltv_data[cc-1]] = value
                  end
                  if r >= 73 && r <= 74 && cc == 10
                    secondary_key = get_value value
                    @jumbo_rate_adjustment[primary_key][secondary_key] = {}
                  end
                  if r >= 73 && r <= 74 && cc >= 15 && cc <= 16
                    @jumbo_rate_adjustment[primary_key][secondary_key][@ltv_data[cc-1]] = {}
                    @jumbo_rate_adjustment[primary_key][secondary_key][@ltv_data[cc-1]] =  value
                  end
                  # MISCELLANEOUS
                  if r == 77 && cc == 10
                    m_key = value
                    @jumbo_other_adjustment[primary_key][m_key] = {}
                  end
                  if r == 77 && cc == 16
                    @jumbo_other_adjustment[primary_key][m_key] = value
                  end
                end
              end
              (0..3).each do |cc|
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "MAX PRICE AFTER ADJUSTMENTS"
                    max_key = "LoanType/LA/"
                    @jumbo_other_adjustment[max_key] = {}
                  end
                  if r >= 71 && r <= 72 && cc == 1
                    key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
                    @jumbo_other_adjustment[max_key][key] = {}
                  end
                  if r >= 71 && r <= 72 && cc == 3
                    @jumbo_other_adjustment[max_key][key] = value
                  end
                end
              end
            end
          end
          adjustment = [@jumbo_purchase_adjustment,@jumbo_rate_adjustment,@jumbo_other_adjustment]
          make_adjust(adjustment,sheet)

          create_program_association_with_adjustment(sheet)
        end
      end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def jumbo_6600
    @programs_ids = []
    @purchase_adjustment = {}
    @rate_adjustment = {}
    @adjustment_hash = {}
    @other_adjustment = {}
    primary_key = ''
    secondary_key = ''
    cltv_key = ''
    key = ''
    adj_key = ''
    @xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6600")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        (10..35).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              program_property
              @programs_ids << @program.id
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # adjustments
        (39..84).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(40)
          @max_data = sheet_data.row(82)
          if row.compact.count >= 1
            (0..14).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Purchase Transaction"
                  primary_key = "LoanPurpose/FICO/LTV"
                  @purchase_adjustment[primary_key] = {}
                elsif value == "Rate/Term Transaction"
                  primary_key = "LoanType/Term/FICO/LTV"
                  @rate_adjustment[primary_key] = {}
                elsif value == "Cash Out Transaction"
                  primary_key = "LoanPurpose/RefinanceOption/LTV"
                  @adjustment_hash[primary_key] = {}
                elsif value == "MISCELLANEOUS"
                  primary_key = "Miscellaneous"
                  @other_adjustment[primary_key] = {}
                elsif value == "MAX PRICE AFTER ADJUSTMENTS"
                  primary_key = "LoanType/LA/"
                  @other_adjustment[primary_key] = {}
                end
                # Purchase Transaction Adjustment
                if r >= 41 && r <= 47 && cc == 1
                  secondary_key = get_value value
                  @purchase_adjustment[primary_key][secondary_key] = {}
                end
                if r >= 41 && r <= 47 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @purchase_adjustment[primary_key][secondary_key][cltv_key] = {}
                  @purchase_adjustment[primary_key][secondary_key][cltv_key] = value
                end

                # Rate/Term Transaction Adjustment
                if r >= 50 && r <= 56 && cc == 1
                  secondary_key = get_value value
                  @rate_adjustment[primary_key][secondary_key] = {}
                end
                if r >= 50 && r <= 56 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @rate_adjustment[primary_key][secondary_key][cltv_key] = {}
                  @rate_adjustment[primary_key][secondary_key][cltv_key] = value
                end

                # Cash Out Transaction Adjustment
                if r >= 59 && r <= 74 && cc == 1
                  if value.include?("Loan Amount")
                    secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                  else
                    secondary_key = get_value value
                  end
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 59 && r <= 74 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                end
                if r == 75 && cc == 1
                  secondary_key = "Escrow-Waiver/NY"
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r == 75 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                end
                if r == 76 && cc == 1
                  secondary_key = "FL"
                  adj_key = "NV"
                  @adjustment_hash[primary_key][secondary_key] = {}
                  @adjustment_hash[primary_key][adj_key] = {}
                end
                if r == 76 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][adj_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                  @adjustment_hash[primary_key][adj_key][cltv_key] = value
                end

                # Other Adjustments
                if r == 79 && cc == 1
                  secondary_key = "NY"
                  @other_adjustment[primary_key][secondary_key] = {}
                end
                if r == 79 && cc == 4
                  @other_adjustment[primary_key][secondary_key] = value
                end
                if r >= 83 && r <= 84 && cc == 1
                  key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
                  @other_adjustment[primary_key][key] = {}
                end
                if r >= 83 && r <= 84 && cc >=2 && cc <= 4
                  @other_adjustment[primary_key][key][@max_data[cc-1]] = {}
                  @other_adjustment[primary_key][key][@max_data[cc-1]] = value
                end
              end
            end
          end
        end
        adjustment = [@purchase_adjustment,@rate_adjustment,@adjustment_hash,@other_adjustment]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
      end
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def jumbo_7600
    @programs_ids = []
    @purchase_adjustment = {}
    @rate_adjustment = {}
    @adjustment_hash = {}
    @other_adjustment = {}
    primary_key = ''
    secondary_key = ''
    cltv_key = ''
    key = ''
    adj_key = ''
    @xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 7600")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        (10..36).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              program_property
              @programs_ids << @program.id
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # adjustments
        (40..85).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(41)
          @max_data = sheet_data.row(83)
          if row.compact.count >= 1
            (0..14).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Purchase Transaction"
                  primary_key = "LoanPurpose/FICO/LTV"
                  @purchase_adjustment[primary_key] = {}
                elsif value == "Rate/Term Transaction"
                  primary_key = "LoanType/Term/FICO/LTV"
                  @rate_adjustment[primary_key] = {}
                elsif value == "Cash Out Transaction"
                  primary_key = "LoanPurpose/RefinanceOption/LTV"
                  @adjustment_hash[primary_key] = {}
                elsif value == "MISCELLANEOUS"
                  primary_key = "Miscellaneous"
                  @other_adjustment[primary_key] = {}
                elsif value == "MAX PRICE AFTER ADJUSTMENTS"
                  primary_key = "LoanType/LA/"
                  @other_adjustment[primary_key] = {}
                end
                # Purchase Transaction Adjustment
                if r >= 42 && r <= 48 && cc == 1
                  secondary_key = get_value value
                  @purchase_adjustment[primary_key][secondary_key] = {}
                end
                if r >= 42 && r <= 48 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @purchase_adjustment[primary_key][secondary_key][cltv_key] = {}
                  @purchase_adjustment[primary_key][secondary_key][cltv_key] = value
                end

                # Rate/Term Transaction Adjustment
                if r >= 51 && r <= 57 && cc == 1
                  secondary_key = get_value value
                  @rate_adjustment[primary_key][secondary_key] = {}
                end
                if r >= 51 && r <= 57 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @rate_adjustment[primary_key][secondary_key][cltv_key] = {}
                  @rate_adjustment[primary_key][secondary_key][cltv_key] = value
                end

                # Cash Out Transaction Adjustment
                if r >= 60 && r <= 75 && cc == 1
                  if value.include?("Loan Amount")
                    secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                  else
                    secondary_key = get_value value
                  end
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 60 && r <= 75 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                end
                if r == 76 && cc == 1
                  secondary_key = "Escrow-Waiver/NY"
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r == 76 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                end
                if r == 77 && cc == 1
                  secondary_key = "FL"
                  adj_key = "NV"
                  @adjustment_hash[primary_key][secondary_key] = {}
                  @adjustment_hash[primary_key][adj_key] = {}
                end
                if r == 77 && cc >= 6 && cc <= 14
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][adj_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                  @adjustment_hash[primary_key][adj_key][cltv_key] = value
                end

                # Other Adjustments
                if r == 80 && cc == 1
                  secondary_key = "NY"
                  @other_adjustment[primary_key][secondary_key] = {}
                end
                if r == 80 && cc == 4
                  @other_adjustment[primary_key][secondary_key] = value
                end
                if r >= 84 && r <= 85 && cc == 1
                  key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
                  @other_adjustment[primary_key][key] = {}
                end
                if r >= 84 && r <= 85 && cc >=2 && cc <= 4
                  @other_adjustment[primary_key][key][@max_data[cc-1]] = {}
                  @other_adjustment[primary_key][key][@max_data[cc-1]] = value
                end
              end
            end
          end
        end
        adjustment = [@purchase_adjustment,@rate_adjustment,@adjustment_hash,@other_adjustment]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
      end
    end
    # redirect_to programs_import_file_path(@bank)
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end


  def jumbo_6400
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6400")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @flex_hash = {}
        @jumbo_flex_hash = {}
        primary_key = ''
        secondary_key = ''
        @cltv_data = []
        (10..41).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              if @title.present? && cc < 9
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all

                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
              end
              if @block_hash.keys.first == "Rate"
                if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        (44..58).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              if cc < 5 && @title == "10/1 ARM - 6410"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
              end
              if @title.present? && cc < 9 && @title == "10/1 ARM - 6410"
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
                # @program.adjustments.destroy_all

                @block_hash = {}
                key = ''
                  (1..50).each do |max_row|
                    @data = []
                    (0..3).each_with_index do |index, c_i|
                      rrr = rr + max_row -1
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        elsif (c_i == 1)
                          @block_hash[key][21] = value
                        elsif (c_i == 2)
                          @block_hash[key][30] = value
                        elsif (c_i == 3)
                          @block_hash[key][45] = value
                        end
                        @data << value
                      end
                    end
                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                end
                if @block_hash.keys.first == "Rate"
                  if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        #Adjustments
        (10..19).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(13)
          if row.compact.count >= 1
            (10..16).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "FLEX JUMBO 6400 SERIES ADJUSTMENTS"
                  primary_key = "Jumbo/LoanType/FICO/LTV"
                  @flex_hash[primary_key] = {}
                end
                if r >= 14 && r <= 19 && cc == 10
                  secondary_key = get_value value
                  @flex_hash[primary_key][secondary_key] = {}
                end
                if r >= 14 && r <= 19 && cc >= 12 && cc <= 16
                  cltv_key = get_value @cltv_data[cc-1]
                  @flex_hash[primary_key][secondary_key][cltv_key] = {}
                  @flex_hash[primary_key][secondary_key][cltv_key] = value
                end
              end
            end
          end
        end
        adjustment = [@flex_hash]
        make_adjust(adjustment,sheet)

        (21..38).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(13)
          if row.compact.count >= 1
            (10..16).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "FLEX JUMBO 6400 SERIES ADJUSTMENTS"
                  primary_key = "Jumbo/LoanType/FICO/LTV"
                  @jumbo_flex_hash[primary_key] = {}
                end
                if r >= 23 && r <= 38 && cc == 10
                  if value.include?("Loan Amount")
                    secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                  elsif value.include?("Cash Out")
                    secondary_key = "Cashout/Fico/ltv"
                  else
                    secondary_key = get_value value
                  end
                  @jumbo_flex_hash[primary_key][secondary_key] = {}
                end
                if r >= 23 && r <= 38 && cc == 16
                  @jumbo_flex_hash[primary_key][secondary_key] = value
                end
              end
            end
          end
        end
        adjustment = [@jumbo_flex_hash]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
      end
    end
    # redirect_to programs_import_file_path(@bank)
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end
  def jumbo_6800
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6800")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        primary_key = ''
        first_key = ''
        cltv_key = ''
        c_val = ''
        @block_adjustment = {}
        @misc_adjustment = {}
        (10..37).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
              @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
              program_property
              @programs_ids << @program.id
              # @program.adjustments.destroy_all
              @block_hash = {}
              key = ''
              (1..50).each do |max_row|
                @data = []
                (0..3).each_with_index do |index, c_i|
                  rrr = rr + max_row -1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    elsif (c_i == 1)
                      @block_hash[key][21] = value
                    elsif (c_i == 2)
                      @block_hash[key][30] = value
                    elsif (c_i == 3)
                      @block_hash[key][45] = value
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        #Adjustment
        (40..50).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(42)
          if (row.compact.count >= 1)
            #Higher of LTV/CLTV Adjustment
            (0..11).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "PRIME JUMBO 6800 SERIES ADJUSTMENTS"
                  primary_key = "Jumbo/LoanPurpose/FICO/LTV"
                  @block_adjustment[primary_key] = {}
                end

                if r >= 43 && r <= 50 && cc == 1
                  cltv_key = get_value value
                  @block_adjustment[primary_key][cltv_key] = {}
                end

                if r >= 43 && r <= 50 && cc >= 4 && cc <= 11
                  key_val = get_value @key_data[cc-1]
                  @block_adjustment[primary_key][cltv_key][key_val] = value
                end
              end
            end

            #MISCELLANEOUS Adjustment
            (13..16).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "MISCELLANEOUS"
                  first_key = "Miscellaneous"
                  @misc_adjustment[first_key] = {}
                end

                if r >= 43 && r <= 44 && cc == 13
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @misc_adjustment[first_key][value] = c_val
                end
              end
            end
          end
        end
        adjustment = [@misc_adjustment,@block_adjustment]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
      end

    end
    # redirect_to programs_import_file_path(@bank)
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def jumbo_6900_7900
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "JUMBO 6900 & 7900")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @cltv_data = []
        @other_data = []
        @jumbo_data = []
        @jumbo_other_data = []
        @adjustment_hash = {}
        @other_adjustment = {}
        @jumbo_adjustment_hash = {}
        @jumbo_other_adjustment = {}
        primary_key = ''
        secondary_key = ''
        max_key = ''
        key = ''
        cltv_key = ''
        (10..23).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
              if @title.present?
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
              end
              if @block_hash.values.first.values.first == "21 Day"
                if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        (51..64).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 1

              @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                program_property
                @programs_ids << @program.id
              if @title.present?
                # @program.adjustments.destroy_all
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row -1
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      elsif (c_i == 1)
                        @block_hash[key][21] = value
                      elsif (c_i == 2)
                        @block_hash[key][30] = value
                      elsif (c_i == 3)
                        @block_hash[key][45] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
              end
              if @block_hash.values.first.values.first == "21 Day"
                if @block_hash.values.first.keys.first.nil?
                @block_hash.values.first.shift
              end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Adjustment
        (26..44).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(28)
          if row.compact.count >= 1
            (1..11).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "RENEW JUMBO QM 6900 SERIES ADJUSTMENTS"
                  primary_key = "RateType/LTV/FICO"
                  @adjustment_hash[primary_key] = {}
                end

                # Purchase Transaction Adjustment
                if r >= 29 && r <= 44 && cc == 1
                  if value.include?("Loan Amount")
                    secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                  elsif value.include?("Cashout")
                    secondary_key = "Cashout/Fico/ltv"
                  elsif value.include?("Condo")
                    secondary_key = "Condo"
                  elsif value.include?("Escrow")
                    secondary_key = "Escrow Waiver/NY"
                  else
                    secondary_key = get_value value
                  end
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 29 && r <= 44 && cc >= 5 && cc <= 11
                  cltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][cltv_key] = value
                end
              end
            end

            #other adjustment
            (13..16).each do |cc|
              value = sheet_data.cell(r,cc)
              @other_data = sheet_data.row(r) if r == 32
              if value.present?
                if value == "MAX PRICE AFTER ADJUSTMENTS"
                  max_key = "RateType/LA/"
                  @other_adjustment[max_key] = {}
                end
                if r >= 33 && r <= 34 && cc == 13
                  key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
                  @other_adjustment[max_key][key] = {}
                end
                if r >= 33 && r <= 34 && cc >= 14 && cc <= 16
                  other_key = get_value @other_data[cc-1]
                  @other_adjustment[max_key][key][other_key] = {}
                  @other_adjustment[max_key][key][other_key] = value
                end
              end
            end

          end
        end
        adjustment = [@adjustment_hash,@other_adjustment]
        make_adjust(adjustment,sheet)

        #second adjustment
        (67..85).each do |r|
          row = sheet_data.row(r)
          @jumbo_data = sheet_data.row(69)
          if row.compact.count >= 1
            (1..11).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "RENEW JUMBO NON-QM 7900 SERIES ADJUSTMENTS"
                  primary_key = "RateType/LTV/FICO"
                  @jumbo_adjustment_hash[primary_key] = {}
                end

                # Purchase Transaction Adjustment
                if r >= 70 && r <= 85 && cc == 1
                  if value.include?("Loan Amount")
                    secondary_key = value.include?("<") ? "0"+value.split("Loan Amount").last : value.split("Loan Amount").last
                  elsif value.include?("Cashout")
                    secondary_key = "Cashout/Fico/ltv"
                  elsif value.include?("Condo")
                    secondary_key = "Condo"
                  elsif value.include?("Escrow")
                    secondary_key = "Escrow Waiver/NY"
                  else
                    secondary_key = get_value value
                  end
                  @jumbo_adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 70 && r <= 85 && cc >= 5 && cc <= 11
                  cltv_key = get_value @jumbo_data[cc-1]
                  @jumbo_adjustment_hash[primary_key][secondary_key][cltv_key] = {}
                  @jumbo_adjustment_hash[primary_key][secondary_key][cltv_key] = value
                end
              end
            end

            #other adjustment
            (13..16).each do |cc|
              value = sheet_data.cell(r,cc)
              @jumbo_other_data = sheet_data.row(r) if r == 73
              if value.present?
                if value == "MAX PRICE AFTER ADJUSTMENTS"
                  max_key = "RateType/LA/"
                  @jumbo_other_adjustment[max_key] = {}
                end
                if r >=74 && r <= 75 && cc == 13
                  key = value.include?("LA <") ? "0" + value.split("LA").last : value.split("LA").last
                  @jumbo_other_adjustment[max_key][key] = {}
                end
                if r >= 74 && r <= 75 && cc >= 14 && cc <= 16
                  other_key = get_value @jumbo_other_data[cc-1]
                  @jumbo_other_adjustment[max_key][key][other_key] = {}
                  @jumbo_other_adjustment[max_key][key][other_key] = value
                end
              end
            end

          end
        end
        adjustment = [@jumbo_adjustment_hash,@jumbo_other_adjustment]
        make_adjust(adjustment,sheet)

        create_program_association_with_adjustment(sheet)
      end
    end
    # redirect_to programs_import_file_path(@bank)
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def get_value value1
    if value1.present?
      if value1.include?("FICO <")
        value1 = "0"+value1.split("FICO").last
      elsif value1.include?("<")
        value1 = "0"+value1
      elsif value1.include?("FICO")
        value1 = value1.split("FICO ").last.first(9)
      elsif value1 == "Investment Property"
        value1 = "Property/Type"
      else
        value1
      end
    end
  end

  def get_key value1
    if value1.present?
      if value1.include?("Streamline")
        value1 = "FHA/Refinance"
      elsif value1.include?("CashOut")
        value1 = "VA/CashOut"
      elsif value1.include?("Lock")
        value1 = "FHA/Refinance"
      elsif value1.include?("IRRR")
        value1 = "FHA/Refinance"
      end
    end
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  def create_program_association_with_adjustment(sheet)
    adjustment_list = Adjustment.where(sheet_name: sheet)
    program_list = Program.where(sheet_name: sheet)

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

  private
    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def read_sheet
      file = File.join(Rails.root,  'OB_CMG_Wholesale7575.xls')
      @xlsx = Roo::Spreadsheet.open(file)
    end

    def get_program
      @program = Program.find(params[:id])
    end

    def get_key value1
      if value1.present?
        if value1.include?("Streamline")
          value1 = "FHA/Refinance"
        elsif value1.include?("CashOut")
          value1 = "VA/CashOut"
        elsif value1.include?("Lock")
          value1 = "FHA/Refinance"
        elsif value1.include?("IRRR")
          value1 = "FHA/Refinance"
        end
      end
    end

    def program_property
      # term
      if (@program.program_name.split("Year").count > 1) 
        term = @program.program_name.split("Year").first.tr('^0-9><%', '')
      elsif (@program.program_name.split("Yr").count > 1)
        term = @program.program_name.split("Yr").first.tr('^0-9><%', '')  
      end
      # Arm Basic
      if @program.program_name.split("ARM").count > 1
        arm_basic = @program.program_name.split("ARM").first.split("/").first
      end
        # loan type
      if @program.program_name.include?("Fixed")
        loan_type = "Fixed"
      elsif @program.program_name.include?("ARM")
        loan_type = "ARM"
      elsif @program.program_name.include?("Floating")
        loan_type = "Floating"
      elsif @program.program_name.include?("Variable")
        loan_type = "Variable"
      else
        loan_type = nil
      end
      # streamline
      streamline = false
      fha = false
      va = false
      usda = false
      full_doc = false
      if @program.program_name.include?("FHA")
        streamline = true
        fha = true
        full_doc = true
      elsif @program.program_name.include?("VA")
        streamline = true
        va = true
        full_doc = true
      elsif @program.program_name.include?("USDA")
        streamline = true
        usda = true
        full_doc = true
      end
      # High Balance
      jumbo_high_balance = false
      if @program.program_name.include?("High Bal") || @program.program_name.include?("HIGH BAL")
        jumbo_high_balance = true
      end
      # Program Property  
      if @program.program_name.split("-").count > 1
        program_category = @program.program_name.split("-").last
      end
         # Loan Limit Type
      if @program.program_name.include?("Non-Conforming")
        @program.loan_limit_type << "Non-Conforming"
      end
      if @program.program_name.include?("Conforming")
        @program.loan_limit_type << "Conforming"
      end
      if @program.program_name.include?("Jumbo")
        @program.loan_limit_type << "Jumbo"
      end
      if @program.program_name.include?("High Balance")
        @program.loan_limit_type << "High Balance"
      end
      @program.save 
      @program.update(term: term,loan_type: loan_type,loan_purpose: "Purchase",program_category: program_category, jumbo_high_balance: jumbo_high_balance, streamline: streamline,fha: fha, va: va, usda: usda, full_doc: full_doc, arm_basic: arm_basic)
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        Adjustment.create(data: hash,sheet_name: sheet)
      end
    end
end
