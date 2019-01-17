class ObAlliedMortgageGroupWholesale8570Controller < ApplicationController
  # before_action :get_sheet, only: [:programs, :ak]
  # before_action :get_program, only: [:single_program]
  def index
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "Cover")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Allied Mortgage"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def fha
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "FHA")
        sheet_data = xlsx.sheet(sheet)
        @fha_adjustment = {}
        @indi_adjustment = {}
        @lock_adjustment = {}
        @avrg_adjustment = {}
        @alhs_adjustment = {}
        @key_data = []
        @key_data2 = []
        f_key = ''
        first_key = ''
        cc = ''
        ccc = ''
        c_val = ''
        # Adjustments FHA
        (39..57).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "FHA & USDA Loan Level Adjustments"
                  first_key = "FHA/USDA Loan"
                  @fha_adjustment[first_key] = {}
                end

                if r >=41 && r <= 52 && cc == 17
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @fha_adjustment[first_key][value] = c_val
                end

                if r >=56 && r <= 57 && cc == 17
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @fha_adjustment[first_key][value] = c_val
                end

              end
            end
          end
        end

        # Adjustments INDICES
        (84..94).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (1..10).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "INDICES"
                  f_key = "INDICES"
                  @indi_adjustment[f_key] = {}
                end

                if r >=85 && r <= 89 && cc == 2
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @indi_adjustment[f_key][value] = c_val
                end

                if value == "Lock Expiration Dates:"
                  first_key = "Lock/ExpirationDates:"
                  @lock_adjustment[first_key] = {}
                end

                if r >=85 && r <= 87 && cc == 7
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @lock_adjustment[first_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Avrage Price
        (92..94).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (2..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Average Price Offer Rate (APOR) For Loans Locked During The Week Of 01/14/2019"
                  f_key = "AveragePrice"
                  @avrg_adjustment[f_key] = {}
                end

                if r == 93 && cc >= 2 && cc <= 9
                  rr = r+1
                  c_val = sheet_data.cell(rr,cc)
                  @avrg_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Allied Holesale
        (84..95).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj"
                  f_key = "AlliedWholesale/LoanAmount/adj"
                  @alhs_adjustment[f_key] = {}
                end

                if r >= 87 && r <= 95 && cc == 17
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @alhs_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end
        Adjustment.create(data: @fha_adjustment, sheet_name: sheet)
        Adjustment.create(data: @indi_adjustment, sheet_name: sheet)
        Adjustment.create(data: @lock_adjustment, sheet_name: sheet)
        Adjustment.create(data: @avrg_adjustment, sheet_name: sheet)
        Adjustment.create(data: @alhs_adjustment, sheet_name: sheet)
      end
    end
    redirect_to dashboard_index_path
  end

  def va
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "VA")
        sheet_data = xlsx.sheet(sheet)
        @valn_adjustment = {}
        @awlm_adjustment = {}
        @indi_adjustment = {}
        @avrg_adjustment = {}
        @lock_adjustment = {}
        primary_key = ''
        second_key = ''
        c_val = ''
        f_key = ''

        # VA loan level adjustment
        (15..33).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "VA Loan Level Adjustments"
                  primary_key = "VA/RateType/FICO/LTV"
                  @valn_adjustment[primary_key] = {}
                end

                if r >=17 && r <= 31 && cc == 17
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end

                if r >=32 && r <= 33 && cc == 17
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end
              end
            end
          end
        end

        # WholeSale adjustment
        (34..45).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (17..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj *"
                  primary_key = "VA/RateType/FICO/LTV"
                  @awlm_adjustment[primary_key] = {}
                end

                if r >=37 && r <= 45 && cc == 17
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @awlm_adjustment[primary_key][value] = c_val
                end
              end
            end
          end
        end

        # INDICES adjustment
        (49..54).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (18..20).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "INDICES"
                  primary_key = "VA/RateType/FICO/LTV"
                  second_key = "INDICES"
                  @indi_adjustment[primary_key] = {}
                  @indi_adjustment[primary_key][second_key] = {}
                end

                if r >=50 && r <= 54 && cc == 18
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @indi_adjustment[primary_key][second_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Avrage Price
        (88..90).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (2..9).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Average Price Offer Rate (APOR) For Loans Locked During The Week Of 01/14/2019"
                  f_key = "AveragePrice"
                  @avrg_adjustment[f_key] = {}
                end

                if r == 89 && cc >= 2 && cc <= 9
                  rr = r+1
                  c_val = sheet_data.cell(rr,cc)
                  @avrg_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustment lock Expiration
        (83..86).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (2..5).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Lock Expiration Dates:"
                  f_key = "Lock/ExpirationDates:"
                  @lock_adjustment[f_key] = {}
                end

                if r >=84 && r <= 86 && cc == 2
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @lock_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end
        Adjustment.create(data: @valn_adjustment, sheet_name: sheet)
        Adjustment.create(data: @awlm_adjustment, sheet_name: sheet)
        Adjustment.create(data: @indi_adjustment, sheet_name: sheet)
        Adjustment.create(data: @lock_adjustment, sheet_name: sheet)
      end
    end
    redirect_to dashboard_index_path
  end

  def conf_fixed
    file = File.join(Rails.root,  'OB_Allied_Mortgage_Group_Wholesale8570.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "CONF FIXED")
        sheet_data = xlsx.sheet(sheet)
        @valn_adjustment = {}
        @avrg_adjustment = {}
        @lock_adjustment = {}
        primary_key = ''
        second_key = ''
        c_val = ''
        f_key = ''

        # VA loan level adjustment
        (88..102).each do |r|
          row = sheet_data.row(r)
          if (row.compact.count >= 1)
            (13..16).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Allied Wholesale Loan Amt Adj"
                  primary_key = "AlliedWholesale/LoanAmount/"
                  @valn_adjustment[primary_key] = {}
                end

                if r >=91 && r <= 99 && cc == 13
                  value = get_value value
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end

                if r >= 101 && r <= 102 && cc == 13
                  value = get_value value
                  ccc = cc + 2
                  c_val = sheet_data.cell(r,ccc)
                  @valn_adjustment[primary_key][value] = c_val
                end
              end
            end
          end
        end

        #Adjustments Avrage Price
        (109..111).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (7..14).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Average Price Offer Rate (APOR) For Loans Locked During The Week Of 01/14/2019"
                  f_key = "AveragePrice"
                  @avrg_adjustment[f_key] = {}
                end

                if r == 110 && cc >= 7 && cc <= 14
                  rr = r+1
                  c_val = sheet_data.cell(rr,cc)
                  @avrg_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end

        # #Adjustment lock Expiration
        (109..112).each do |r|
          row = sheet_data.row(r)
          @key_data = sheet_data.row(93)
          @key_data2 = sheet_data.row(93)
          if (row.compact.count >= 1)
            (2..5).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Lock Expiration Dates:"
                  f_key = "Lock/ExpirationDates:"
                  @lock_adjustment[f_key] = {}
                end

                if r >=110 && r <= 112 && cc == 2
                  ccc = cc + 3
                  c_val = sheet_data.cell(r,ccc)
                  @lock_adjustment[f_key][value] = c_val
                end
              end
            end
          end
        end
        Adjustment.create(data: @valn_adjustment, sheet_name: sheet)
        Adjustment.create(data: @avrg_adjustment, sheet_name: sheet)
        Adjustment.create(data: @lock_adjustment, sheet_name: sheet)
      end
    end
    redirect_to dashboard_index_path
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private

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

    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_program
      @program = Program.find(params[:id])
    end
end

