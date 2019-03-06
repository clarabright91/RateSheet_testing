class ObRoyalPacificFundingWholesale8409Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :royal_pfc]
  before_action :get_program, only: [:single_program, :program_property]
  before_action :read_sheet, only: [:index,:royal_pfc]

  def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "Royal PFC")
          @name = "Royal Pacific Funding"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
    end
  end

  def royal_pfc
    @xlsx.sheets.each do |sheet|
      if (sheet == "Royal PFC")
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (104..620).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count > 1) && (row.compact.count <= 4))
            rr = r + 2
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4*max_column + 2
              @title = sheet_data.cell(r,cc)
              begin
                if @title.present? && @title.length > 8 && !@title.include?("Programs") && !@title.include?("Adjusters") && !@title.include?("Amount") && !@title.include?("Margin") && !@title.include?("Max") && !@title.include?("$") && !@title.include?("FICO") && !@title.include?("not allowed") && @title != "HomeReady Mortgage Caps" && !@title.include?("Non-Escrowed")
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @programs_ids << @program.id
                  program_property sheet
                  @block_hash = {}
                  key = ''
                  (1..14).each do |max_row|
                    @data = []
                    (0..2).each_with_index do |index, c_i|
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        # (694..708).each do |r|
        #   row = sheet_data.row(r)
        #   row = row.reject { |e| e.to_s.empty? }
        #   if ((row.compact.count > 1) && (row.compact.count <= 3))
        #     rr = r + 2
        #     max_column_section = row.compact.count - 1
        #     (0..max_column_section).each do |max_column|
        #       cc = 4*max_column + 2
        #       cc = 7 if (max_column == 1)
        #       @title = sheet_data.cell(r,cc)
        #       begin
        #         if @title.present? && @title.length > 8 && !@title.include?("Programs") && !@title.include?("Adjusters") && !@title.include?("Amount") && !@title.include?("Margin") && !@title.include?("Max") && !@title.include?("$") && !@title.include?("FICO") && !@title.include?("not allowed") && @title != "HomeReady Mortgage Caps" && !@title.include?("Non-Escrowed")
        #           @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
        #           @programs_ids << @program.id
        #           program_property sheet
        #           @block_hash = {}
        #           key = ''
        #           (1..14).each do |max_row|
        #             @data = []
        #             # column_count = 5 if (cc == 9)
        #             (0..2).each_with_index do |index, c_i|
        #               rrr = rr + max_row
        #               ccc = cc + c_i
        #               value = sheet_data.cell(rrr,ccc)
        #               if value.present?
        #                 if (c_i == 0)
        #                   key = value
        #                   @block_hash[key] = {}
        #                 else
        #                   @block_hash[key][15*c_i] = value if key.present?
        #                 end
        #                 @data << value
        #               end
        #             end
        #             if @data.compact.reject { |c| c.blank? }.length == 0
        #               break # terminate the loop
        #             end
        #           end
        #           @program.update(base_rate: @block_hash)
        #         end
        #       rescue Exception => e
        #         error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, sheet_name: sheet, error_detail: e.message)
        #         error_log.save
        #       end
        #     end
        #   end
        # end
      end
    end
    redirect_to programs_ob_royal_pacific_funding_wholesale8409_path(@sheet_obj)
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

  def program_property value1
    # term
    if @program.program_name.include?("30 Year") || @program.program_name.include?("30Yr") || @program.program_name.include?("30 Yr") || @program.program_name.include?("30/25 Year")
      term = 30
    elsif @program.program_name.include?("20 Year") || @program.program_name.include?("20 Yr")
      term = 20
    elsif @program.program_name.include?("15 Year") || @program.program_name.include?("15 Yr")
      term = 15
    elsif @program.program_name.include?("10 Year") || @program.program_name.include?("10 Yr")
      term = 10
    else
      term = nil
    end

    # Arm Advanced
    if @program.program_name.include?("5/2/5")
      arm_advanced = "5/2/5"
    end

    # Loan Size
    if @title.include?("HIGH BAL") || @title.include?("HIGH BALANCE") || @title.include?("High Balance")
      loan_size = "High-Balance"
      jumbo_high_balance = true
    elsif @title.include?("Conforming")
      loan_size = "Conforming"
      conforming = true
    elsif @title.include?("Jumbo")
      loan_size = "Jumbo"
    end

    # Arm Basic
    if @program.program_name.include?("1/1")
      arm_basic = 1
    elsif @program.program_name.include?("2/1")
      arm_basic = 2
    elsif @program.program_name.include?("3/1")
      arm_basic = 3
    elsif @program.program_name.include?("5/1")
      arm_basic = 5
    elsif @program.program_name.include?("7/1")
        arm_basic = 7
    elsif @program.program_name.include?("10/1")
      arm_basic = 10
    end

    # Loan-Type
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
    if @program.program_name.include?("High Balance") || @program.program_name.include?("High Bal")
      jumbo_high_balance = true
    end

    #fannie_mae_product
    if @program.program_name.include?("HomeReady")
      @program.fannie_mae_product = "HomeReady"
    end

    # fannie-mae and freddie-mac
    if @title.include?("DU")
      fannie_mae = true
    elsif @title.include?("LP")
      freddie_mac = true
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
    @program.update(arm_basic: arm_basic,term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline, jumbo_high_balance: jumbo_high_balance, loan_size: loan_size, sheet_name: @sheet_name, arm_advanced: arm_advanced)
  end

  def get_sheet
    @sheet_obj = Sheet.find(params[:id])
  end

  def read_sheet
    file = File.join(Rails.root,  'OB_Royal_Pacific_Funding_Wholesale8409.xls')
    @xlsx = Roo::Spreadsheet.open(file)
  end
end
