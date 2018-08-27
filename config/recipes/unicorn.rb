set_default(:unicorn_user) { user }
set_default(:unicorn_pid) { "#{current_path}/tmp/pids/unicorn.pid" }
set_default(:unicorn_config) { "#{shared_path}/config/unicorn.rb" }
set_default(:unicorn_log) { "#{shared_path}/log/unicorn.log" }
set_default(:unicorn_workers, 2)

namespace :unicorn do
  desc "Setup Unicorn initializer and app configuration"
  task :setup, roles: :app do
    run "mkdir -p #{shared_path}/config"
    template "unicorn.rb.erb", unicorn_config
    template "unicorn_init.sh.erb", "/tmp/unicorn_init"
    run "chmod +x /tmp/unicorn_init"
    run "sudo mv /tmp/unicorn_init /etc/init.d/unicorn_#{application}"
    run "sudo update-rc.d -f unicorn_#{application} defaults"
  end
  after "deploy:setup", "unicorn:setup"


  # %w[start stop restart].each do |command|
  # desc "#{command} unicorn"
  # task command, roles: :app do
  #    run "service unicorn_#{application} #{command}"
  # end
  # after "deploy:#{command}", "unicorn:#{command}"
  # end

  desc "start unicorn"
  task :start, :roles => :app, :except => {:no_release => true} do
    run "cd #{current_path} && bundle exec unicorn -c #{unicorn_config} -E #{rails_env} -D"
  end
  desc "stop unicorn"
  task :stop, :roles => :app, :except => {:no_release => true} do
    run "sudo kill `cat #{unicorn_pid}`" if remote_file_exists?(unicorn_pid)
  end
  desc "unicorn reload"
  task :reload, :roles => :app, :except => {:no_release => true} do
    run "sudo kill -s USR2 `cat #{unicorn_pid}`"
  end
  desc "graceful stop unicorn"
  task :graceful_stop, :roles => :app, :except => {:no_release => true} do
    run "sudo kill -s QUIT `cat #{unicorn_pid}`"
  end
  desc "restart unicorn"
  task :restart, :roles => :app, :except => {:no_release => true} do
    stop
    start
  end

  after "deploy:restart", "unicorn:restart"
end
