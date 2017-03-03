after "deploy:stop",    "delayed_job:stop"
after "deploy:start",   "delayed_job:start"
after "deploy:restart", "delayed_job:restart"

namespace :delayed_job do
  desc "start delayed_job"
  task :start, :roles => :app, :except => {:no_release => true} do
    run "cd #{current_path} &&  RAILS_ENV=#{rails_env} bin/delayed_job start"
  end

  desc "stop delayed_job"
  task :stop, :roles => :app, :except => {:no_release => true} do
    run "cd #{current_path} &&  RAILS_ENV=#{rails_env} bin/delayed_job stop"
  end

  desc "restart delayed_job"
  task :restart, :roles => :app, :except => {:no_release => true} do
    stop
    start
  end
end