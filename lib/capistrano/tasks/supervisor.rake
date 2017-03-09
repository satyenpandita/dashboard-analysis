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
        execute :sudo, "supervisorctl update"
        execute :sudo, "supervisorctl start #{fetch(:application)}"
    end
  end

end