require 'test_helper'

class ImportFilesControllerTest < ActionDispatch::IntegrationTest
  test "should get index" do
    get import_files_index_url
    assert_response :success
  end

  test "should get new" do
    get import_files_new_url
    assert_response :success
  end

  test "should get create" do
    get import_files_create_url
    assert_response :success
  end

end
