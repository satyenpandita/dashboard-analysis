namespace :nginx do
  desc "Install latest stable release of nginx"
  task :install do
    on roles(:all) do |host|
        execute :sudo, "apt-get -y update"
        execute :sudo, "apt-get -y install nginx"
    end
  end

  %w[start stop restart].each do |command|
    desc "#{command} nginx"
    task command do
        on roles(:all) do |host|
            execute :sudo, "service nginx #{command}"
        end
    end
  end

  desc "Setup nginx configuration for this application"
  task :setup do
    on roles(:all) do |host|
        template "nginx.erb", "/tmp/nginx_conf"
        execute :sudo, "mv /tmp/nginx_conf /etc/nginx/sites-enabled/#{fetch(:application)}"
        execute :sudo, "rm -f /etc/nginx/sites-enabled/default"
    end
  end


end
