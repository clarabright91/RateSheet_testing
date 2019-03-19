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
  
  def get_conf
    conf = ["CONF HB"]
    # conforming += Acronym.new(conforming).to_a
    conf += conf.map(&:downcase)
    return conf
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
    loan_size = get_conforming + get_non_conforming + get_high_balance + get_jumbo + get_conf
    return loan_size
  end

  def update_fields p_name
    set_loan_size(p_name)           if fetch_loan_size_fields.each{ |word| p_name.downcase.include?(word.downcase) }
    set_load_type(p_name)           if ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_fha                         if p_name.downcase.include?("fha")
    set_va                          if p_name.downcase.include?("va")
    set_usda                        if p_name.downcase.include?("usda")
    set_loan_purpose(p_name)        if ["Purchase", "Refinance"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_arm_basic(p_name)           if ["ARM"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_arm_advanced(p_name)        if ["ARM"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_fannie_mae(p_name)          if ["Fannie Mae", "DU"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_freddie_mac(p_name)         if ["Freddie Mac", "LP"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_freddie_mac_product(p_name) if ["Home Possible"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_term(p_name) if (5..50).to_a.collect{|n| n.to_s}.each{ |word| p_name.downcase.include?(word.downcase) }
    self.save
  end

  def set_term(prog_name)
    self.term = ProgramUpdate.set_term(prog_name)
  end

  def set_load_type(prog_name)
    present_word = nil
    ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].each{ |word|
      present_word = word if prog_name.downcase.include?(word.downcase)
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
    ["Purchase", "Refinance"].each{ |word|
      present_word = word if p_name.downcase.include?(word.downcase)
    }
    self.loan_purpose = present_word
  end

  def set_arm_basic p_name
    self.arm_basic = ProgramUpdate.arm_basic(p_name)
  end

  def set_loan_size p_name
    present_word = nil
    fetch_loan_size_fields.each{ |word|
      present_word = word if p_name.downcase.include?(word.downcase)
    }
    loan_size = get_high_balance.include?(present_word) ? "High-Balance" : get_jumbo.include?(present_word) ? "Jumbo" : get_conforming.include?(present_word) ? "Conforming" : get_non_conforming.include?(present_word) ? "Non-Conforming" : get_conf.include?(present_word) ? "Conforming and High-Balance" : nil
    self.loan_size = loan_size
  end

  def set_arm_advanced p_name
    present_word = nil
    ["10/5", "5/1 3-2-5"].each{ |word|
      present_word = word if p_name.downcase.include?(word.downcase)
    }
    self.arm_advanced = present_word
  end

  def set_fannie_mae p_name
    present_word = nil
    ["Fannie Mae", "DU"].each{ |word|
      present_word = true if p_name.downcase.include?(word.downcase)
    }
    self.fannie_mae = present_word
  end

  def set_freddie_mac p_name
    present_word = nil
    ["Freddie Mac", "LP"].each{ |word|
      present_word = true if p_name.downcase.include?(word.downcase)
    }
    self.freddie_mac = present_word
  end

  def set_freddie_mac_product p_name
    present_word = nil
    ["Home Possible"].each{ |word|
      present_word = word if p_name.downcase.include?(word.downcase)
    }
    self.freddie_mac_product = present_word
  end
end
