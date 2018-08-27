namespace :redis do

  desc "Install the latest release of Redis"
  task :install do
    on roles(:all) do |host|
        execute :sudo, "apt-get -y install redis-server"
    end
  end

  # Just to use if you need to do more than the default configuration, mind to change the monit script details as well
  # desc "Setup Redis"
  # task :setup do
  #   run "#{sudo} cp /etc/redis/redis.conf /etc/redis/redis.conf.default"
  #   template "redis.conf.erb", "/tmp/redis.conf"
  #   run "#{sudo} mv /tmp/redis.conf /etc/redis/redis.conf"
  #   restart
  # end
  # after "deploy:setup", "redis:setup"

  %w[start stop restart].each do |command|
    desc "#{command} redis"
    task command do
        on roles(:all) do |host|
            execute :sudo, "service redis-server #{command}"
        end
    end
  end
end