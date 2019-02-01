class HardWorker
  include Sidekiq::Worker

  def perform(id)
    debugger
  end
end
