# This file is auto-generated from the current state of the database. Instead
# of editing this file, please use the migrations feature of Active Record to
# incrementally modify your database, and then regenerate this schema definition.
#
# Note that this schema.rb definition is the authoritative source for your
# database schema. If you need to create the application database on another
# system, you should be using db:schema:load, not running all the migrations
# from scratch. The latter is a flawed and unsustainable approach (the more migrations
# you'll amass, the slower it'll run and the greater likelihood for issues).
#
# It's strongly recommended that you check this file into your version control system.

ActiveRecord::Schema.define(version: 2018_12_28_121535) do

  # These are extensions that must be enabled in order to support this database
  enable_extension "plpgsql"

  create_table "adjustments", force: :cascade do |t|
    t.json "data"
    t.string "program_title"
    t.string "sheet_name"
    t.integer "program_ids", default: [], array: true
    t.integer "program_id"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "banks", force: :cascade do |t|
    t.string "name"
    t.integer "nmls"
    t.string "phone"
    t.string "address1"
    t.string "address2"
    t.string "city"
    t.string "state_code"
    t.string "zip"
    t.string "state_eligibility"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "program_adjustments", force: :cascade do |t|
    t.integer "program_id"
    t.integer "adjustment_id"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "programs", force: :cascade do |t|
    t.integer "bank_id"
    t.string "title"
    t.integer "loan_type"
    t.integer "term"
    t.integer "interest_type"
    t.integer "interest_subtype"
    t.boolean "jumbo_high_balance"
    t.boolean "conforming"
    t.boolean "fannie_mae"
    t.boolean "fannie_mae_home_ready"
    t.boolean "freddie_mac"
    t.boolean "freddie_mac_home_possible"
    t.boolean "fha"
    t.boolean "va"
    t.boolean "usda"
    t.boolean "streamline"
    t.boolean "full_doc"
    t.text "adjustments"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
    t.string "sheet_name"
    t.json "base_rate"
  end

end
