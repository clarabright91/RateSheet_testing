class ObUnionHomeMortgageWholesale1711Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :conventional, :conven_highbalance_30, :gov_highbalance_30, :government_30_15_yr, :arm_programs, :fnma_du_refi_plus, :fhlmc_open_access, :fnma_home_ready, :fhlmc_home_possible, :simple_access, :jumbo_fixed]
  before_action :read_sheet, only: [:index, :conventional, :conven_highbalance_30, :gov_highbalance_30, :government_30_15_yr, :arm_programs, :fnma_du_refi_plus, :fhlmc_open_access, :fnma_home_ready, :fhlmc_home_possible, :simple_access, :jumbo_fixed]
  before_action :get_program, only: [:single_program, :program_property]

  def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "Intro")
          # headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Union Home Mortgage Wholesale"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def conventional
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conventional")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @mortgage_hash = {}
        @sub_hash = {}
        @property_hash = {}
        @multiunit_hash = {}
        primary_key = ''
        secondary_key = ''
        new_key = ''
        ltv_key = ''
        # programs
        (10..77).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..16).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # adjustments
        (79..144).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(80)
          @cltv_data = sheet_data.row(90)
          @sub_data = sheet_data.row(104)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == 'Cash-Out Refinance'
                  primary_key = "RefinanceOption/FICO/LTV"
                  new_key = "Cash out"
                  @adjustment_hash[primary_key] = {}
                  @adjustment_hash[primary_key][new_key] = {}
                end
                if value == "Applicable for all mortgages with terms greater than 15 years  "
                  primary_key = "Term/FICO/LTV"
                  new_key = "15-Inf"
                  @mortgage_hash[primary_key] = {}
                  @mortgage_hash[primary_key][new_key] = {}
                end
                if value == "Subordinate Financing LTV/CLTV/FICO Adjustments"
                  primary_key = "FinancingType/LTV/CLTV/FICO"
                  new_key = "Subordinate Financing"
                  @sub_hash[primary_key] = {}
                  @sub_hash[primary_key][new_key] = {}
                end
                # Cash-Out Refinance
                if r >= 81 && r <= 87 && cc == 3
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][new_key][secondary_key] = {}
                end
                if r >= 81 && r <= 87 && cc >= 4 && cc <= 7
                  ltv_key = get_value @ltv_data[cc-1]
                  if ltv_key.include?("%")
                    ltv_key = ltv_key.split("%").last
                  end
                  @adjustment_hash[primary_key][new_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][new_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with all terms
                if r == 91 && cc == 3
                  primary_key = "PropertyType/Term"
                  @adjustment_hash[primary_key] = {}
                end
                if r == 91 && cc >= 4 && cc <= 11
                  ltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with terms greater than 15 years
                if r >= 95 && r <= 101 && cc == 3
                  secondary_key = get_value value
                  @mortgage_hash[primary_key][new_key][secondary_key] = {}
                end
                if r >= 95 && r <= 101 && cc >= 4 && cc <= 11
                  ltv_key = get_value @cltv_data[cc-1]
                  if ltv_key.include?('%')
                    ltv_key = ltv_key.split("%").last
                  end
                  @mortgage_hash[primary_key][new_key][secondary_key][ltv_key] = {}
                  @mortgage_hash[primary_key][new_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Financing LTV/CLTV/FICO Adjustments
                if r >= 105 && r <= 109 && cc == 3
                  secondary_key = get_value value
                  if secondary_key.include?("%")
                    secondary_key = secondary_key.split("%").first+secondary_key.split("%").last
                  end
                  @sub_hash[primary_key][new_key][secondary_key] = {}
                end
                if r >= 105 && r <= 109 && cc == 4
                  ltv_key = get_value value
                  if ltv_key.include?("%")
                    ltv_key = ltv_key.split("%").first+ltv_key.split("%").last
                  end
                  @sub_hash[primary_key][new_key][secondary_key][ltv_key] = {}
                end
                if r >= 105 && r <= 109 && cc >= 5 && cc <= 6
                  sub_data = get_value @sub_data[cc-1]
                  @sub_hash[primary_key][new_key][secondary_key][ltv_key][sub_data] = {}
                  @sub_hash[primary_key][new_key][secondary_key][ltv_key][sub_data] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Finance
                if r == 112 && cc == 3
                  primary_key = "FinancingType"
                  @sub_hash[primary_key] = value
                end
                # Property Type
                if r == 112 && cc == 6
                  primary_key = "PropertyType/Term/LTV"
                  secondary_key = "Condo"
                  ltv_key = "15-Inf"
                  new_key = "75-Inf"
                  @property_hash[primary_key] = {}
                  @property_hash[primary_key][secondary_key] = {}
                  @property_hash[primary_key][secondary_key][ltv_key] = {}
                  cc = cc + 3
                  new_value = sheet_data.cell(r,cc)
                  @property_hash[primary_key][secondary_key][ltv_key][new_key] = new_value
                end
                if value == "Multi Unit Property"
                  primary_key = "PropertyType/LTV"
                  @multiunit_hash[primary_key] = {}
                end
                if r >= 112 && r <= 113 && cc == 11
                  new_key = "2 Unit" if value.include?("2 unit")
                  new_key = "3-4 Unit" if value.include?("3-4 Unit")
                  ltv_key = "0-85" if value.include?("<=85")
                  ltv_key = "0-75" if value.include?("<=75")
                  @multiunit_hash[primary_key] = {}
                  @multiunit_hash[primary_key][new_key] = {}
                  cc = cc + 2
                  new_value = sheet_data.cell(r,cc)
                  @multiunit_hash[primary_key][new_key][ltv_key] = new_value
                end
                if r == 113 && cc == 6
                  primary_key = "PropertyType"
                  new_key = "Manufactured Home"
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key][new_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key][new_key] = new_value
                  end
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@mortgage_hash,@sub_hash,@property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def conven_highbalance_30
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conven HighBalance 30")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (11..22).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..10).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def gov_highbalance_30
    @xlsx.sheets.each do |sheet|
      if (sheet == "GOV HighBalance 30")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (10..19).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..8).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def government_30_15_yr
    @xlsx.sheets.each do |sheet|
      if (sheet == "Government 30_15 Yr")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @property_hash = {}
        @fico_hash ={}
        primary_key = ''
        second_key = ''
        new_key = true
        # programs
        (11..42).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..16).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # Adjustments
        (30..77).each do |r|
          row = sheet_data.row(r)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Government Price Adjustments"
                  @property_hash["PropertyType"]={}
                  @fico_hash["FICO"] = {}
                end
                if r >= 31 && r <= 33 && cc == 11
                  primary_key = "VA" if value == "VA"
                  primary_key = "USDA" if value == "USDA"
                  primary_key = "FHA" if value == "FHA Streamline"
                  @adjustment_hash[primary_key] = {}
                  @adjustment_hash[primary_key][new_key] = {}
                  if @adjustment_hash[primary_key][new_key] = {}
                    cc = cc + 1
                    new_value = sheet_data.cell(r,cc)
                    @adjustment_hash[primary_key][new_key] = new_value
                  end
                end
                if r == 34 && cc == 11
                  @adjustment_hash["VA/RefinanceOption"] = {}
                  @adjustment_hash["VA/RefinanceOption"]["true"] = {}
                  cc = cc + 1
                  new_value = sheet_data.cell(r,cc)
                  @adjustment_hash["VA/RefinanceOption"]["true"]["IRRRL"] =new_value
                end
                if r == 35 && cc == 11
                  @adjustment_hash["FICO"] = {}
                  # cc = cc + 1
                  new_value = sheet_data.cell(r,cc+1)
                  @adjustment_hash["FICO"]["720-Inf"]=new_value
                end
                if r >= 36 && r <= 39 && cc == 11
                  second_key = value.split(" ").last
                  new_value = sheet_data.cell(r,cc+1)
                  @fico_hash["FICO"][second_key]=new_value
                end
                if r >= 41 && r <= 42 && cc == 11
                  new_value = sheet_data.cell(r,cc+1)
                  @property_hash["PropertyType"][value]=new_value
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@fico_hash, @property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def arm_programs
    @xlsx.sheets.each do |sheet|
      if (sheet == "ARM Programs")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @mortgage_hash = {}
        @sub_hash = {}
        @property_hash = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        # programs
        (11..55).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..13).each do |max_row|
                  @data = []
                  (0..3).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*(c_i+1)] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # adjustments
        (58..123).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(59)
          @cltv_data = sheet_data.row(68)
          @sub_data = sheet_data.row(82)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == 'Cash-Out Refinance'
                  primary_key = "RefinanceOption/FICO/LTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Applicable for all mortgages with terms greater than 15 years  "
                  primary_key = "Term/FICO/LTV"
                  @mortgage_hash[primary_key] = {}
                end
                if value == "Subordinate Financing LTV/CLTV/FICO Adjustments"
                  primary_key = "FinancingType/LTV/CLTV/FICO"
                  @sub_hash[primary_key] = {}
                end
                # Cash-Out Refinance
                if r >= 60 && r <= 65 && cc == 2
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 60 && r <= 65 && cc >= 3 && cc <= 6
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with all terms
                if r == 69 && cc == 2
                  primary_key = "PropertyType/Term"
                  @adjustment_hash[primary_key] = {}
                end
                if r == 69 && cc >= 3 && cc <= 10
                  ltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with terms greater than 15 years
                if r >= 73 && r <= 79 && cc == 2
                  secondary_key = get_value value
                  @mortgage_hash[primary_key][secondary_key] = {}
                end
                if r >= 73 && r <= 79 && cc >= 3 && cc <= 10
                  ltv_key = get_value @cltv_data[cc-1]
                  @mortgage_hash[primary_key][secondary_key][ltv_key] = {}
                  @mortgage_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Financing LTV/CLTV/FICO Adjustments
                if r >= 83 && r <= 87 && cc == 5
                  secondary_key = get_value value
                  @sub_hash[primary_key][secondary_key] = {}
                end
                if r >= 83 && r <= 87 && cc == 6
                  ltv_key = get_value value
                  @sub_hash[primary_key][secondary_key][ltv_key] = {}
                end
                if r >= 83 && r <= 87 && cc >= 7 && cc <= 8
                  sub_data = get_value @sub_data[cc-1]
                  @sub_hash[primary_key][secondary_key][ltv_key][sub_data] = {}
                  @sub_hash[primary_key][secondary_key][ltv_key][sub_data] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Finance
                if r == 90 && cc == 2
                  primary_key = "FinancingType"
                  @sub_hash[primary_key] = value
                end
                # Property Type
                if r == 90 && cc == 5
                  primary_key = "Condo/LTV/Term/>75%"
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 90 && cc == 9
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 1
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 91 && cc == 5
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 91 && cc == 9
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 1
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@mortgage_hash,@sub_hash,@property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def fnma_du_refi_plus
    @xlsx.sheets.each do |sheet|
      if (sheet == "FNMA DU-Refi Plus")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @mortgage_hash = {}
        @sub_hash = {}
        @property_hash = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        # programs
        (10..24).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 4
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..13).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # adjustments
        (27..98).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(35)
          @cltv_data = sheet_data.row(45)
          @sub_data = sheet_data.row(59)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == 'Cash-Out Refinance'
                  primary_key = "RefinanceOption/FICO/LTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Applicable for all mortgages with terms greater than 15 years  "
                  primary_key = "Term/FICO/LTV"
                  @mortgage_hash[primary_key] = {}
                end
                if value == "Subordinate Financing LTV/CLTV/FICO Adjustments"
                  primary_key = "FinancingType/LTV/CLTV/FICO"
                  @sub_hash[primary_key] = {}
                end
                # Cash-Out Refinance
                if r >= 36 && r <= 42 && cc == 2
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 36 && r <= 42 && cc >= 3 && cc <= 6
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with all terms
                if r == 46 && cc == 2
                  primary_key = "PropertyType/Term"
                  @adjustment_hash[primary_key] = {}
                end
                if r == 46 && cc >= 3 && cc <= 10
                  ltv_key = get_value @cltv_data[cc-1]
                  @adjustment_hash[primary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with terms greater than 15 years
                if r >= 50 && r <= 56 && cc == 2
                  secondary_key = get_value value
                  @mortgage_hash[primary_key][secondary_key] = {}
                end
                if r >= 50 && r <= 56 && cc >= 3 && cc <= 10
                  ltv_key = get_value @cltv_data[cc-1]
                  @mortgage_hash[primary_key][secondary_key][ltv_key] = {}
                  @mortgage_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Financing LTV/CLTV/FICO Adjustments
                if r >= 60 && r <= 64 && cc == 2
                  secondary_key = get_value value
                  @sub_hash[primary_key][secondary_key] = {}
                end
                if r >= 60 && r <= 64 && cc == 3
                  ltv_key = get_value value
                  @sub_hash[primary_key][secondary_key][ltv_key] = {}
                end
                if r >= 60 && r <= 64 && cc >= 4 && cc <= 5
                  sub_data = get_value @sub_data[cc-1]
                  @sub_hash[primary_key][secondary_key][ltv_key][sub_data] = {}
                  @sub_hash[primary_key][secondary_key][ltv_key][sub_data] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Finance
                if r == 67 && cc == 2
                  primary_key = "FinancingType"
                  @sub_hash[primary_key] = value
                end
                # Property Type
                if r == 67 && cc == 5
                  primary_key = "Condo/LTV/Term/>75%"
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 67 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 68 && cc == 5
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 68 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@mortgage_hash,@sub_hash,@property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def fhlmc_open_access
    @xlsx.sheets.each do |sheet|
      if (sheet == "FHLMC Open Access")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @mortgage_hash = {}
        @sub_hash = {}
        @property_hash = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        # programs
        (11..25).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..13).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # Adjustments
        (27..91).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(28)
          @cltv_data = sheet_data.row(39)
          @sub_data = sheet_data.row(51)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == 'Cash-Out Refinance'
                  primary_key = "RefinanceOption/FICO/LTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Applicable for all mortgages with terms greater than 15 years  "
                  primary_key = "Term/FICO/LTV"
                  @mortgage_hash[primary_key] = {}
                end
                if value == "Subordinate Financing LTV/CLTV/FICO Adjustments"
                  primary_key = "FinancingType/LTV/CLTV/FICO"
                  @sub_hash[primary_key] = {}
                end
                # Cash-Out Refinance
                if r >= 29 && r <= 35 && cc == 2
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 29 && r <= 35 && cc >= 3 && cc <= 6
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with terms greater than 15 years
                if r >= 40 && r <= 48 && cc == 2
                  secondary_key = get_value value
                  @mortgage_hash[primary_key][secondary_key] = {}
                end
                if r >= 40 && r <= 48 && cc >= 3 && cc <= 10
                  ltv_key = get_value @cltv_data[cc-1]
                  @mortgage_hash[primary_key][secondary_key][ltv_key] = {}
                  @mortgage_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Financing LTV/CLTV/FICO Adjustments
                if r >= 52 && r <= 58 && cc == 6
                  secondary_key = get_value value
                  @sub_hash[primary_key][secondary_key] = {}
                end
                if r >= 52 && r <= 58 && cc == 7
                  ltv_key = get_value value
                  @sub_hash[primary_key][secondary_key][ltv_key] = {}
                end
                if r >= 52 && r <= 58 && cc >= 8 && cc <= 9
                  sub_data = get_value @sub_data[cc-1]
                  @sub_hash[primary_key][secondary_key][ltv_key][sub_data] = {}
                  @sub_hash[primary_key][secondary_key][ltv_key][sub_data] = (value.class == Float ? value*100 : value)
                end
                # Property Type
                if r == 61 && cc == 5
                  primary_key = "Condo/LTV/Term/>75%"
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 61 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 62 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@mortgage_hash,@sub_hash,@property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def fnma_home_ready
    @xlsx.sheets.each do |sheet|
      if (sheet == "FNMA Home Ready")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @sub_hash = {}
        @property_hash = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        # programs
        (11..28).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 4
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..16).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # Adjustment
        (35..90).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(36)
          @cltv_data = sheet_data.row(50)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Applicable for all mortgages with with all terms"
                  primary_key = "PropertyType/LTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Applicable for all mortgages with terms greater than 15 years  "
                  primary_key = "Term/FICO/LTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Subordinate Financing LTV/CLTV/FICO Adjustments"
                  primary_key = "FinancingType/LTV/CLTV/FICO"
                  @sub_hash[primary_key] = {}
                end
                # Applicable for all mortgages with with all terms
                if r == 37 && cc >= 3 && cc <= 10
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Applicable for all mortgages with terms greater than 15 years
                if r >= 41 && r <= 47 && cc == 2
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 41 && r <= 47 && cc >= 3 && cc <= 10
                  ltv_key = get_value @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Financing LTV/CLTV/FICO Adjustments
                if r >= 51 && r <= 55 && cc == 5
                  secondary_key = get_value value
                  @sub_hash[primary_key][secondary_key] = {}
                end
                if r >= 51 && r <= 55 && cc == 6
                  ltv_key = get_value value
                  @sub_hash[primary_key][secondary_key][ltv_key] = {}
                end
                if r >= 51 && r <= 55 && cc >= 7 && cc <= 8
                  cltv_key = get_value @cltv_data[cc-1]
                  @sub_hash[primary_key][secondary_key][ltv_key][cltv_key] = {}
                  @sub_hash[primary_key][secondary_key][ltv_key][cltv_key] = (value.class == Float ? value*100 : value)
                end
                 # Property Type
                 if r == 59 && cc == 5
                  primary_key = "Condo/LTV/Term/>75%"
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 59 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 60 && cc == 5
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 60 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@sub_hash,@property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def fhlmc_home_possible
    @xlsx.sheets.each do |sheet|
      if (sheet == "FHLMC Home Possible")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @sub_hash = {}
        @property_hash = {}
        primary_key = ''
        secondary_key = ''
        ltv_key = ''
        # programs
        (10..27).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 4
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..16).each do |max_row|
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
                        @block_hash[key][15*c_i] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # Adjustment
        (35..86).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(36)
          @cltv_data = sheet_data.row(46)
          if row.compact.count >= 1
            (0..13).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Applicable for all mortgages with terms greater than 15 years  "
                  primary_key = "Term/FICO/LTV"
                  @adjustment_hash[primary_key] = {}
                end
                if value == "Subordinate Financing LTV/CLTV/FICO Adjustments"
                  primary_key = "FinancingType/LTV/CLTV/FICO"
                  @sub_hash[primary_key] = {}
                end
                # Applicable for all mortgages with terms greater than 15 years
                if r >= 37 && r <= 43 && cc == 2
                  secondary_key = get_value value
                  @adjustment_hash[primary_key][secondary_key] = {}
                end
                if r >= 37 && r <= 43 && cc >= 3 && cc <= 10
                  ltv_key = @ltv_data[cc-1]
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = {}
                  @adjustment_hash[primary_key][secondary_key][ltv_key] = (value.class == Float ? value*100 : value)
                end
                # Subordinate Financing LTV/CLTV/FICO Adjustments
                if r >= 47 && r <= 51 && cc == 5
                  secondary_key = value
                  @sub_hash[primary_key][secondary_key] = {}
                end
                if r >= 47 && r <= 51 && cc == 6
                  ltv_key = value
                  @sub_hash[primary_key][secondary_key][ltv_key] = {}
                end
                if r >= 47 && r <= 51 && cc >= 7 && cc <= 8
                  cltv_key = @cltv_data[cc-1]
                  @sub_hash[primary_key][secondary_key][ltv_key][cltv_key] = {}
                  @sub_hash[primary_key][secondary_key][ltv_key][cltv_key] = (value.class == Float ? value*100 : value)
                end
                 # Property Type
                 if r == 55 && cc == 5
                  primary_key = "Condo/LTV/Term/>75%"
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 55 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 56 && cc == 5
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 3
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
                if r == 56 && cc == 10
                  primary_key = value
                  @property_hash[primary_key] = {}
                  if @property_hash[primary_key] = {}
                    cc = cc + 2
                    new_value = sheet_data.cell(r,cc)
                    @property_hash[primary_key] = new_value
                  end
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@sub_hash,@property_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def simple_access
    @xlsx.sheets.each do |sheet|
      if (sheet == "Simple Access")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (9..83).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 5
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
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
                        @block_hash[key][15*(c_i-1)+5] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # Adjustment
        (90..122).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(93)
          @cltv_data = sheet_data.row(96)
          if row.compact.count >= 1
            (0..15).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "OTHER LLPAs (Price Adjustments)(1)(2)"
                  first_key = "FinancingType/LTV/CLTV/FICO"
                  @other_hash[first_key] = {}
                end
                if r >= 94 && r <= 105 && cc == 10
                  ltv_key = get_key value
                  @other_hash[first_key][ltv_key] = {}
                end
                if r >= 94 && r <= 105 && cc >= 11 && cc <= 15
                  third_key = @ltv_data[cc-3]
                  @other_hash[first_key][ltv_key][third_key] = value
                end
                if value == "FICO/LTV LLPAs (Price Adjustments)"
                  secondary_key = "FICO/LTV"
                  @adjustment_hash[secondary_key] = {}
                end

                if r >= 97 && r <= 105 && cc == 3
                  ltv_key = value
                  @adjustment_hash[secondary_key][ltv_key] = {}
                end
                if r >= 97 && r <= 105 && cc >= 4 && cc <= 8
                  third_key = @cltv_data[cc-3]
                  @adjustment_hash[secondary_key][ltv_key][third_key] = value
                end

                if value == "Product LLPAs (Price Adjustments)"
                  secondary_key = "Product LLPAs"
                  @llpa_hash[secondary_key] = {}
                end
                if r >= 91 && r <= 93 && cc == 3
                  ltv_key = value
                  cltv_key = sheet_data.cell(r,cc+1)
                  @llpa_hash[secondary_key][ltv_key] = cltv_key
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@other_hash,@llpa_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def jumbo_fixed
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Fixed")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (9..32).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if (row.compact.count <= 1)
            rr = r + 1
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = max_column + 7
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                program_property sheet
                @block_hash = {}
                key = ''
                (1..10).each do |max_row|
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
                        @block_hash[key][15*(c_i-1)+5] = value if key.present?
                      end
                      @data << value
                    end
                  end
                  if @data.compact.reject { |c| c.blank? }.length == 0
                    break # terminate the loop
                  end
                end
                @program.update(base_rate: @block_hash)
              end
            end
          end
        end
        # Adjustment
        (36..63).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(38)
          if row.compact.count >= 1
            (0..15).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Other Adjustments"
                  first_key = "FinancingType/LTV/CLTV/FICO"
                  @other_hash[first_key] = {}
                end
                if value == "State"
                  second_key = "State"
                  @other_hash[first_key][second_key] = {}
                end
                if value == "Property Type"
                  second_key = "Property Type"
                  @other_hash[first_key][second_key] = {}
                end
                if value == "2nd Home (700+)"
                  second_key = "2nd Home (700+)"
                  @other_hash[first_key][second_key] = {}
                end
                if value == "Miscellaneous"
                  second_key = "Miscellaneous"
                  @other_hash[first_key][second_key] = {}
                end
                if value == "Purpose"
                  second_key = "Purpose"
                  @other_hash[first_key][second_key] = {}
                end

                if r >= 39 && r <= 42 && cc == 10
                  ltv_key = value
                  @other_hash[first_key][second_key][ltv_key] = {}
                end

                if r >= 39 && r <= 42 && cc >= 11 && cc <= 15
                  third_key = @ltv_data[cc-3]
                  @other_hash[first_key][second_key][ltv_key][third_key] = value
                end

                if r >= 45 && r <= 46 && cc == 10
                  ltv_key = value
                  @other_hash[first_key][second_key][ltv_key] = {}
                end

                if r >= 45 && r <= 46 && cc >= 11 && cc <= 15
                  third_key = @ltv_data[cc-3]
                  @other_hash[first_key][second_key][ltv_key][third_key] = value
                end

                if r == 49 && cc == 10
                  ltv_key = value
                  @other_hash[first_key][second_key][ltv_key] = {}
                end

                if r == 49 && cc >= 11 && cc <= 15
                  third_key = @ltv_data[cc-3]
                  @other_hash[first_key][second_key][ltv_key][third_key] = value
                end

                if r >= 52 && r <= 53 && cc == 10
                  ltv_key = value
                  @other_hash[first_key][second_key][ltv_key] = {}
                end

                if r >= 52 && r <= 53 && cc >= 11 && cc <= 15
                  third_key = @ltv_data[cc-3]
                  @other_hash[first_key][second_key][ltv_key][third_key] = value
                end

                if r >= 56 && r <= 57 && cc == 10
                  ltv_key = value
                  @other_hash[first_key][second_key][ltv_key] = {}
                end

                if r >= 56 && r <= 57 && cc >= 11 && cc <= 15
                  third_key = @ltv_data[cc-3]
                  @other_hash[first_key][second_key][ltv_key][third_key] = value
                end

                if value == "LTV/CLTV/HLTV/FICO "
                  first_key = "FinancingType/LTV/CLTV/FICO"
                  @adjustment_hash[first_key] = {}
                end
                if value == "<=1.0MM"
                  secondary_key = "<=1.0MM"
                  @adjustment_hash[first_key][secondary_key] = {}
                end
                if value == "<=1.5MM"
                  secondary_key = "<=1.5MM"
                  @adjustment_hash[first_key][secondary_key] = {}
                end
                if value == "<=2.0MM"
                  secondary_key = "<=2.0MM"
                  @adjustment_hash[first_key][secondary_key] = {}
                end
                if r >= 39 && r <= 43 && cc == 3
                  ltv_key = value
                  @adjustment_hash[first_key][secondary_key][ltv_key] = {}
                end
                if r >= 39 && r <= 43 && cc >= 4 && cc <= 8
                  third_key = @ltv_data[cc-3]
                  @adjustment_hash[first_key][secondary_key][ltv_key][third_key] = value
                end
                if r >= 46 && r <= 49 && cc == 3
                  ltv_key = value
                  @adjustment_hash[first_key][secondary_key][ltv_key] = {}
                end
                if r >= 46 && r <= 49 && cc >= 4 && cc <= 8
                  third_key = @ltv_data[cc-3]
                  @adjustment_hash[first_key][secondary_key][ltv_key][third_key] = value
                end
                if r >= 52 && r <= 55 && cc == 3
                  ltv_key = value
                  @adjustment_hash[first_key][secondary_key][ltv_key] = {}
                end
                if r >= 52 && r <= 55 && cc >= 4 && cc <= 8
                  third_key = @ltv_data[cc-3]
                  @adjustment_hash[first_key][secondary_key][ltv_key][third_key] = value
                end

              end
            end
          end
        end
        adjustment = [@adjustment_hash,@other_hash]
        make_adjust(adjustment,sheet)
      end
    end
    redirect_to programs_ob_union_home_mortgage_wholesale1711_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  def get_program
    @program = Program.find(params[:id])
  end

  private

  def get_sheet
    @sheet_obj = Sheet.find(params[:id])
  end

  def get_value value1
    if value1.present?
      if value1.include?("<=") || value1.include?("<")
        value1 = "0-"+value1.split("<=").last.tr('^0-9 ', '')
      elsif value1.include?(">")
        value1 = value1.split(">").last.tr('^0-9 ', '')+"-Inf"
      else
        value1
      end
    end
  end

  def make_adjust(block_hash, sheet)
    block_hash.each do |hash|
      Adjustment.create(data: hash,sheet_name: sheet)
    end
  end

  def read_sheet
    file = File.join(Rails.root,  'OB_Union_Home_Mortgage_Wholesale1711.xls')
    @xlsx = Roo::Spreadsheet.open(file)
  end

  def program_property sheet
    if @program.program_name.include?("30") || @program.program_name.include?("30/25 Year")
      term = 30
    elsif @program.program_name.include?("20")
      term = 20
    elsif @program.program_name.include?("15")
      term = 15
    elsif @program.program_name.include?("10 Year")
      term = 10
    elsif @program.program_name.include?("5 Year")
      term = 5
    else
      term = nil
    end

      # Loan-Type
      if @program.program_name.include?("Fixed") || @program.program_name.include?("FIXED")
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

      # Streamline Vha, Fha, Usda
      fha = false
      va = false
      usda = false
      streamline = false
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
      if @program.program_name.include?("High Bal") || @program.program_name.include?("High Balance")
        jumbo_high_balance = true
      end

      # Arm Basic
      if @program.program_name.include?("3/1") || @program.program_name.include?("3 / 1")
        arm_basic = 3
      elsif @program.program_name.include?("5/1") || @program.program_name.include?("5 / 1")
        arm_basic = 5
      elsif @program.program_name.include?("7/1") || @program.program_name.include?("7 / 1")
        arm_basic = 7
      elsif @program.program_name.include?("10/1") || @program.program_name.include?("10 / 1")
        arm_basic = 10
      end

      # Arm Advanced
      if @program.program_name.include?("2-2-5 ")
        arm_advanced = "2-2-5"
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
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, jumbo_high_balance: jumbo_high_balance, arm_basic: arm_basic, arm_advanced: arm_advanced, sheet_name: sheet)
    end
  end
