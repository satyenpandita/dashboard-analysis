set_default(:sidekiq_concurrency, 5)

namespace :sidekiq do

  desc 'Generate and upload Upstart configs for daemons needed by the app'
  task :setup_upstart, except: {no_release: true} do
    template 'upstart/sidekiq.conf.erb', '/tmp/sidekiq.conf'
    run "#{sudo} mv /tmp/sidekiq.conf /etc/init/sidekiq.conf"

    template 'upstart/sidekiq-manager.conf.erb', '/tmp/sidekiq-manager.conf'
    run "#{sudo} mv /tmp/sidekiq-manager.conf /etc/init/sidekiq-manager.conf"

    run "#{sudo} touch /etc/sidekiq.conf"
  end
  after 'deploy:install', 'sidekiq:setup_upstart'
  after 'deploy:install', 'sidekiq:manager:start'

  desc "Generate the database.yml configuration file."
  task :setup, roles: :app do
    run "mkdir -p #{shared_path}/config"
    template "sidekiq.yml.erb", "#{shared_path}/config/sidekiq.yml"
  end
  after "deploy:setup", "sidekiq:setup"

  desc "Symlink the database.yml file into latest release"
  task :symlink, roles: :app do
    run "ln -nfs #{shared_path}/config/sidekiq.yml #{release_path}/config/sidekiq.yml"
  end
  after "deploy:finalize_update", "sidekiq:symlink"

  namespace :manager do
    desc 'Start the sidekiq manager'
    task :start do
      sudo 'start sidekiq-manager'
    end

    desc 'Stop the sidekiq manager'
    task :stop do
      sudo 'stop sidekiq-manager || true'
    end

    desc 'Restart the sidekiq manager'
    task :restart do
      stop
      start
    end
  end

  namespace :worker do
    desc 'Start the sidekiq worker'
    task :start do
      sudo "start sidekiq app=#{current_path} index=#{ENV['worker'] || 0}"
    end

    desc 'Stop the sidekiq worker'
    task :stop do
      sudo "stop sidekiq app=#{current_path} index=#{ENV['worker'] || 0}"
    end

    desc 'Restart the sidekiq worker'
    task :restart do
      stop
      start
    end

    desc "Register sidekiq worker with sidekiq manager"
    task :register, roles: :app do
      run "#{sudo} sh -c \"echo #{current_path}, 1 > /etc/sidekiq.conf\""
    end

  end

end


after 'deploy:stop', 'sidekiq:worker:stop'
after 'deploy:start', 'sidekiq:worker:start'
after 'deploy:restart', 'sidekiq:worker:restart'