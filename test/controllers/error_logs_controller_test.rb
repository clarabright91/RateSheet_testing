require 'test_helper'

class ErrorLogsControllerTest < ActionDispatch::IntegrationTest
  test "should get index" do
    get error_logs_index_url
    assert_response :success
  end

end
