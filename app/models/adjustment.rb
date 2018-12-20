class Adjustment < ApplicationRecord
	#validates :program_title, presence: true, uniqueness: {message: "already exists"}
	before_create :check_details

	def check_details
	end
end
