class Program < ApplicationRecord
  belongs_to :bank, optional: true
  belongs_to :sheet, optional: true
  has_many :program_adjustments
  has_many :adjustments, through: :program_adjustments
  belongs_to :sub_sheet, optional: true

  def get_adjustments
    Adjustment.where(sheet_name: self.sheet_name)
  end

  def update_fields p_name
    present_word = nil
    if ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].any? { |word| present_word = word if p_name.include?(word) }
      self.loan_type = present_word
    elsif ["FHA"].any? { |word| p_name.include?(word) }
      self.fha = true
    elsif ["VA"].any? { |word| p_name.include?(word) }
      self.va = true
    elsif ["USDA"].any? { |word| p_name.include?(word) }
      self.usda = true
    elsif ["Purchase", "Refinance"].any? { |word| present_word = word if p_name.include?(word) }
      self.loan_purpose = present_word
    elsif ["1/1", "2/1", "3/1", "7/1", "10/1", "5-1", "5/1"].any? { |word| present_word = word if p_name.include?(word) }
      self.arm_basic = 5 if present_word.eql?("5-1")
      self.arm_basic = present_word.to_i unless present_word.eql?("5-1")
    elsif ["Non-Conforming", "Conforming", "Jumbo", "High-Balance"].any? { |word| present_word = word if p_name.include?(word) }
      self.loan_size = present_word
    elsif ["10/5", "5/1 3-2-5"].any? { |word| present_word = word if p_name.include?(word) }
      self.arm_advanced = present_word
    elsif ["Fannie Mae, DU"].any? { |word| p_name.include?(word) }
      self.fannie_mae = true
    elsif ["Freddie Mac, LP"].any? { |word| p_name.include?(word) }
      self.freddie_mac = true
    elsif ["Home Possible"].any? { |word| present_word = p_name.include?(word) }
      self.freddie_mac_product = present_word
    end
  end
end
