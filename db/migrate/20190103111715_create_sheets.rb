class CreateSheets < ActiveRecord::Migration[5.2]
  def change
    create_table :sheets do |t|
      t.string :name
      t.integer :bank_id

      t.timestamps
    end
  end
end