after "deploy:start",   "memcached:flush"
after "deploy:restart", "memcached:flush"

namespace :memcached do
  desc "Flushes memcached local instance"
  task :flush, :roles => [:app] do
    run "cd #{current_path}; rake memcached:flush RAILS_ENV=#{rails_env}"
  end
end