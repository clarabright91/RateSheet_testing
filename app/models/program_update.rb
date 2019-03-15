class ProgramUpdate
  class << self
    def set_term(name)
      if name.downcase.include?("year")
        digit_string  = name.downcase.split("year").first.split(" ").last
        modify_string(digit_string)
      elsif name.downcase.include?("yr")
        digit_string = name.downcase.split("yr").first.last#.scan(/\d/).join('')
        modify_string(digit_string)
      end
    end

    private
    def modify_string string
      if string.include?("-")
        last_digit   = string.split("-").last.to_i
        digits_count = MatheMatics.digits(last_digit)
        if digits_count > 1
          string = string.gsub("-","").to_i
        else
          string = string.split("-").first + "0" + string.split("-").last
          string = string.to_i
        end
      else
        string.to_i
      end
    end
  end
end
