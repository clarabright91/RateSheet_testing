class ControllerMethodCallingService

	def initialize class_name, action_name
		@class = class_name.gsub("OB", "Ob") if class_name.include?("OB")
		@controllers_names = Dir[Rails.root.join('app/controllers/*_controller.rb')].map { |path| (path.match(/(\w+)_controller.rb/); $1).camelize+"Controller" }
		@action = action_name
	end

	def invoke_method
		# check controller name is exists or not
    if @controllers_names.include?(@class)
      @class = Object.const_get(@class, Class.new(StandardError))
      action = @action.tableize.singularize
      controller = @class.new
      # check methods is present or not in controller
      if controller.methods.include?(@action.downcase.to_sym)
      	sheet = find_sheet(@action)
      	if sheet
      		controller.send(action, sheet.id)
    		else
    			controller.send(:index)
    			sheet = find_sheet(@action)
    			controller.send(action, sheet.id)
    		end
    	end
    end
  end

  private

  def find_sheet sheet_name
  	return Sheet.find_by_name(sheet_name)
  end
end