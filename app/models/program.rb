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

  def update_fields p_name
    set_load_type(p_name)           if ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].any? { |word| p_name.include?(word) }
    set_fha                         if ["FHA"].any? { |word| p_name.include?(word) }
    set_va                          if ["VA"].any? { |word| p_name.include?(word) }
    set_usda                        if ["USDA"].any? { |word| p_name.include?(word) }
    set_loan_purpose(p_name)        if ["Purchase", "Refinance"].any? { |word| p_name.include?(word) }
    set_arm_basic(p_name)           if ["1/1", "2/1", "3/1", "7/1", "10/1", "5-1", "5/1"].any? { |word| p_name.include?(word) }
    set_loan_size(p_name)           if ["Non-Conforming", "Conforming", "Jumbo", "High-Balance"].any? { |word| p_name.include?(word) }
    set_arm_advanced(p_name)        if ["10/5", "5/1 3-2-5"].any? { |word| p_name.include?(word) }
    set_fannie_mae                  if ["Fannie Mae, DU"].any? { |word| p_name.include?(word) }
    set_freddie_mac                 if ["Freddie Mac, LP"].any? { |word| p_name.include?(word) }
    set_freddie_mac_product(p_name) if ["Home Possible"].any? { |word| p_name.include?(word) }
    set_term(p_name) if (5..50).to_a.collect{|n| n.to_s}.any? { |word| p_name.include?(word) }
    self.save
  end


  def set_load_type(prog_name)
    present_word = nil
    ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].any? { |word|
      present_word = word if prog_name.include?(word)
    }
    self.loan_type = present_word
  end

  def set_term(prog_name)
    if prog_name.scan(/\d+/).uniq.count == 1
      self.term = prog_name.scan(/\d+/).uniq[0]
    elsif prog_name.scan(/\d+/).count > 1 && !prog_name.include?("ARM")
      self.term = (prog_name.scan(/\d+/)[0]+ prog_name.scan(/\d+/)[1]).to_i if (MatheMatics.digits(prog_name.scan(/\d+/)[1].to_i).to_i + 1) > 1
      self.term = (prog_name.scan(/\d+/)[0]+ "0" + prog_name.scan(/\d+/)[1]).to_i unless (MatheMatics.digits(prog_name.scan(/\d+/)[1].to_i).to_i + 1) > 1
    end
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
    ["Non-Conforming", "Conforming", "Jumbo", "High-Balance", "High Bal"].any? { |word|
      present_word = word if p_name.include?(word)
    }
    self.loan_size = present_word
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
