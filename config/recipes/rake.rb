namespace :rake_task do
  desc "remote rake task"
  task :default do
    run "cd #{current_path}; bundle exec rake #{ENV['task']} RAILS_ENV=#{rails_env}"
  end
end