set_default(:mysql_encoding, "utf8")
set_default(:mysql_connection_pool, 5)

namespace :mysql do
  # desc "Install the latest stable release of MYSQL."
  # task :install, roles: :db, only: {primary: true} do
  #   run "#{sudo} apt-get -y install libmysqlclient-dev mysql-server",:pty => true do |channel, stream, data|
  #     channel.send_data("#{mysql_root_password}\n\r") if data =~ /password/
  #   end
  # end
  # after "deploy:install", "mysql:install"

  # desc "Create a database for this application."
  # task :create_database, roles: :db, only: {primary: true} do
  #   sql = <<-SQL
  #         CREATE DATABASE IF NOT EXISTS #{mysql_database};
  #         GRANT ALL PRIVILEGES ON #{mysql_database}.* TO #{mysql_user}@#{mysql_host} IDENTIFIED BY '#{mysql_password}';
  #   SQL
  #
  #   run "mysql --user=root -p --execute=\"#{sql}\"" do |channel, stream, data|
  #     if data =~ /^Enter password:/
  #       channel.send_data("#{mysql_root_password}\n\r")
  #     end
  #   end
  # end
  # after "deploy:setup", "mysql:create_database"

  desc "Generate the database.yml configuration file."
  task :setup, roles: :app do
    run "mkdir -p #{shared_path}/config"
    template "mysql.yml.erb", "#{shared_path}/config/database.yml"
  end
  after "deploy:setup", "mysql:setup"

  desc "Symlink the database.yml file into latest release"
  task :symlink, roles: :app do
    run "ln -nfs #{shared_path}/config/database.yml #{release_path}/config/database.yml"
  end
  after "deploy:finalize_update", "mysql:symlink"

end
