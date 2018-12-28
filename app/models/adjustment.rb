class Adjustment < ApplicationRecord
	#validates :program_title, presence: true, uniqueness: {message: "already exists"}
	before_create :check_details
	has_many :program_adjustments
  has_many :programs, through: :program_adjustments
	def check_details
	end
end
