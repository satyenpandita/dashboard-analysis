namespace :supervisor do

  desc "Setup supervisor configuration for this application"
  task :setup do
    on roles(:all) do |host|
        template "supervisor.erb", "/tmp/supervisor_conf"
        execute :sudo, "mv /tmp/supervisor_conf /etc/supervisor/conf.d/#{fetch(:application)}.conf"
        Rake::Task["supervisor:restart"].invoke
    end
  end

  desc "Update Supervisor"
  task :restart do
    on roles(:all) do |host|
        execute :sudo, "systemctl restart supervisor"
    end
  end

  desc "Restart Services"
  task :services_restart do
    on roles(:all) do |host|
        execute :sudo, "supervisorctl restart celery"
        execute :sudo, "supervisorctl restart flower"
        execute :sudo, "supervisorctl restart gunicorn"
    end
  end

end