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
    Adjustment.where(loan_category: self.loan_category)
  end

  def get_non_conforming
    non_conforming = ["Non-Conforming"]
    # non_conforming += Acronym.new(non_conforming).to_a
    non_conforming += non_conforming.map(&:downcase)
    return non_conforming
  end

  def get_conforming
    conforming = ["Conforming","Conf", "FCF"]
    # conforming += Acronym.new(conforming).to_a
    conforming += conforming.map(&:downcase)
    return conforming
  end

  def get_super_conforming
    sup_conf = ["Super Conforming"," SC "]
    sup_conf += sup_conf.map(&:downcase)
    return sup_conf
  end

  def get_non_conf_hb
    non_conf_hb = ["Non-Conforming Jumbo"]
    non_conf_hb += non_conf_hb.map(&:downcase)
    return non_conf_hb
  end
  
  def get_conf
    conf = ["CONF HB","Conf High Bal", "Conforming High Balance"]
    # conforming += Acronym.new(conforming).to_a
    conf += conf.map(&:downcase)
    return conf
  end
  
  def get_high_balance
    high_balance = ["High-Balance", "HIGH BAL", "High Balance", "HB"]
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

  def get_purchase
    lp = ["Purchase","Purch"]
    lp += lp.map(&:downcase)
    return lp
  end

  def get_refinance
    rf = ["Refinance","Refi"]
    rf += rf.map(&:downcase)
    return rf
  end

  def fetch_loan_size_fields
    loan_size = get_conforming + get_non_conforming + get_high_balance + get_jumbo + get_conf + get_non_conf_hb + get_super_conforming
    return loan_size
  end

  def fetch_loan_purpose_fields
    loan_purpose = get_purchase + get_refinance
    return loan_purpose
  end

  def update_fields p_name
    set_loan_size(p_name)           if fetch_loan_size_fields.each{ |word| p_name.downcase.include?(word.downcase) }
    set_load_type(p_name)           if ["Fixed", "ARM", "Hybrid", "Floating", "Variable"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_fha                         if p_name.downcase.include?("fha")
    set_du                          if p_name.downcase.include?("du ")
    set_lp                          if p_name.downcase.include?("lp ")
    set_va                          if p_name.downcase.include?("va")
    set_usda                        if p_name.downcase.include?("usda")
    set_streamline(p_name)          if ["streamline","SL"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_full_doc                    if p_name.downcase.include?("full doc")
    # set_loan_purpose(p_name)        if ["Purchase", "Refinance"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_loan_purpose(p_name)        if fetch_loan_purpose_fields.each{ |word| p_name.downcase.include?(word.downcase) }
    set_conforming(p_name)          if ["Conforming","Conf","fcf"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_fannie_mae(p_name)          if ["Fannie Mae", "FNMA", "Du "].each{ |word| p_name.downcase.include?(word.downcase) }
    set_freddie_mac(p_name)         if ["Freddie Mac", "FHLMC", "Lp "].each{ |word| p_name.downcase.include?(word.downcase) }
    set_freddie_mac_product(p_name) if ["Home Possible","HOME POSSIBLE"].each{ |word| p_name.downcase.include?(word.downcase) }
    set_fannie_mae_product(p_name)  if ["HOMEREADY", "Home Ready"].each{ |word| p_name.downcase.include?(word.downcase) }
    self.save
  end

  def set_load_type(prog_name)
    present_word = nil
    ["Fixed", "ARM", "Hybrid", "Floating", "Variable" ,"FCF"].each{ |word|
      present_word = word if prog_name.downcase.include?(word.downcase)
      if present_word.present? && present_word.downcase.include?('fcf')
        present_word = "Fixed"
      end
    }
    self.loan_type = present_word
  end

  def set_conforming(prog_name)
    present_word = nil
    # self.conforming = true
    ["Conforming","Conf","fcf"].map { |word| 
      present_word = true if prog_name.downcase.include?(word.downcase) 
    }
    self.conforming = present_word 
  end

  def set_du
    self.du = true
  end

  def set_lp
    self.lp = true
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

  def set_streamline(prog_name)
    present_word = nil
    ["streamline","SL"].map { |word| 
      present_word = true if prog_name.downcase.include?(word.downcase) 
    }
    self.streamline = present_word
  end

  def set_full_doc
    self.full_doc = true
  end

  # def set_loan_purpose p_name
  #   present_word = nil
  #   ["Purchase", "Refinance","Purch","Refi"].each{ |word|
  #     debugger
  #     if p_name.downcase.include?(word.downcase)
  #       present_word = word
  #     else
  #       present_word = "Purchase"
  #     end
  #   }
  #   self.loan_purpose = present_word.downcase.capitalize rescue nil
  # end

  def set_loan_purpose p_name
    present_word = nil
    fetch_loan_purpose_fields.each{ |word|
      present_word = word if p_name.downcase.include?(word.downcase)
    }
    loan_purpose = get_refinance.include?(present_word) ? "Refinance" : "Purchase"
    self.loan_purpose = loan_purpose
  end

  def set_arm_basic p_name
    self.arm_basic = ProgramUpdate.arm_basic(p_name)
  end

  def set_loan_size p_name
    present_word = nil
    fetch_loan_size_fields.each{ |word|
      present_word = word if p_name.squish.downcase.include?(word.downcase)
    }
    loan_size = get_high_balance.include?(present_word) ? "High-Balance" : get_jumbo.include?(present_word) ? "Jumbo" : get_super_conforming.include?(present_word) ? "Super Conforming" : get_non_conforming.include?(present_word) ? "Non-Conforming" : get_conforming.include?(present_word) ? "Conforming" : get_conf.include?(present_word) ? "Conforming and High-Balance" : get_non_conf_hb.include?(present_word) ? "Non-Conforming and Jumbo" : nil
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
    present_word = false
    ["Fannie Mae", "FNMA", "DU "].each{ |word|
      present_word = true if p_name.downcase.include?(word.downcase)
    }
    self.fannie_mae = present_word
  end

  def set_freddie_mac p_name
    present_word = false
    ["Freddie Mac", "FHLMC", "LP "].each{ |word|
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

  def set_fannie_mae_product p_name
    present_word = nil
    ["HomeReady", "Home Ready"].each{ |word|
      present_word = "HomeReady" if p_name.downcase.include?(word.downcase)
    }
    self.fannie_mae_product = present_word
  end
end
