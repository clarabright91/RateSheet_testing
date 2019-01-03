require 'test_helper'

class ObCmgWholesalesControllerTest < ActionDispatch::IntegrationTest
  test "should get index" do
    get ob_cmg_wholesales_index_url
    assert_response :success
  end

end
