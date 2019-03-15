class Program < ApplicationRecord
  belongs_to :bank, optional: true
  belongs_to :sheet, optional: true
  has_many :program_adjustments
  has_many :adjustments, through: :program_adjustments
  belongs_to :sub_sheet, optional: true
  before_save :add_bank_name

  def add_bank_name
    self.bank_name = self.sheet.bank.name if self.sheet.present?
    self.bank_name = self.sub_sheet.sheet.bank.name if self.sub_sheet.present?
  end

  def get_adjustments
    Adjustment.where(sheet_name: self.sheet_name)
  end

  def get_non_conforming
    non_conforming = ["Non-Conforming"]
    # non_conforming += Acronym.new(non_conforming).to_a
    non_conforming += non_conforming.map(&:downcase)
    return non_conforming
  end

  def get_conforming
    conforming = ["Conforming"]
    # conforming += Acronym.new(conforming).to_a
    conforming += conforming.map(&:downcase)
    return conforming
  end

  def get_high_balance
    high_balance = ["High-Balance", "HIGH BAL", "High Balance"]
    # high_balance += Acronym.new(high_balance).to_a
    high_balance += high_balance.map(&:downcase)
    return high_balance
  end

  def get_jumbo
    jumbo = ["Jumbo"]
    jumbo += Acronym.new(jumbo).to_a
    jumbo += jumbo.map(&:downcase)
    return jumbo
  end

  def fetch_loan_size_fields
    loan_size = get_conforming + get_non_conforming + get_high_balance + get_jumbo
    return loan_size
  end

  def update_fields p_name
    set_loan_size(p_name)           if fetch_loan_size_fields.any? { |word| p_name.downcase.include?(word.downcase) }
    set_load_type(p_name)           if ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_fha                         if ["FHA"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_va                          if ["VA"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_usda                        if ["USDA"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_loan_purpose(p_name)        if ["Purchase", "Refinance"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_arm_basic(p_name)           if ["1/1", "2/1", "3/1", "7/1", "10/1", "5-1", "5/1"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_arm_advanced(p_name)        if ["10/5", "5/1 3-2-5"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_fannie_mae                  if ["Fannie Mae", "DU"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_freddie_mac                 if ["Freddie Mac", "LP"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_freddie_mac_product(p_name) if ["Home Possible"].any? { |word| p_name.downcase.include?(word.downcase) }
    set_term(p_name) if (5..50).to_a.collect{|n| n.to_s}.any? { |word| p_name.downcase.include?(word.downcase) }
    self.term = ProgramUpdate.set_term(p_name)
    self.save
  end


  def set_load_type(prog_name)
    present_word = nil
    ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].any? { |word|
      present_word = word if prog_name.include?(word)
    }
    self.loan_type = present_word
  end

  def set_fha
    self.fha = true
  end

  def set_va
    self.va = true
  end

  def set_usda
    self.usda = true
  end

  def set_loan_purpose p_name
    present_word = nil
    ["Purchase", "Refinance"].any? { |word|
      present_word = word if p_name.include?(word)
    }
    self.loan_purpose = present_word
  end

  def set_arm_basic p_name
    present_word = nil
    ["1/1", "2/1", "3/1", "7/1", "10/1", "5-1", "5/1"].any? { |word|
      present_word = word if p_name.include?(word)
    }
    self.arm_basic = 5 if present_word.eql?("5-1")
    self.arm_basic = present_word.to_i unless present_word.eql?("5-1")
  end

  def set_loan_size p_name
    present_word = nil
    fetch_loan_size_fields.any? { |word|
      present_word = word if p_name.include?(word)
    }
    loan_size = get_high_balance.include?(present_word) ? "High-Balance" : get_jumbo.include?(present_word) ? "Jumbo" : get_conforming.include?(present_word) ? "Conforming" : "Non-Conforming"
    self.loan_size = loan_size
  end

  def set_arm_advanced p_name
    present_word = nil
    ["10/5", "5/1 3-2-5"].any? { |word|
      present_word = word if p_name.include?(word)
    }
    self.arm_advanced = present_word
  end

  def set_fannie_mae
    self.fannie_mae = true
  end

  def set_freddie_mac
    self.freddie_mac = true
  end

  def set_freddie_mac_product p_name
    present_word = nil
    ["Home Possible"].any? { |word|
      present_word = word if p_name.include?(word)
    }
    self.freddie_mac_product = present_word
  end
end
