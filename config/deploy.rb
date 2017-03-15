# config valid only for current version of Capistrano
lock "3.7.2"

set :application, "auroville"
set :user, "ubuntu"
set :branch, "flask_api"
set :repo_url, "git@github.com:Prashant31/dashboard-analysis.git"
set :pty, true
set :tld, "ec2-35-165-146-82.us-west-2.compute.amazonaws.com"

# Default branch is :master
# ask :flask_api, `git rev-parse --abbrev-ref HEAD`.chomp

# Default deploy_to directory is /var/www/my_app_name
# set :deploy_to, "/var/www/my_app_name"

# Default value for :format is :airbrussh.
# set :format, :airbrussh

# You can configure the Airbrussh format using :format_options.
# These are the defaults.
# set :format_options, command_output: true, log_file: "log/capistrano.log", color: :auto, truncate: :auto

# Default value for :pty is false
set :pty, true

# Default value for :linked_files is []
# append :linked_files, "config/database.yml", "config/secrets.yml"

# Default value for linked_dirs is []
# append :linked_dirs, "log", "tmp/pids", "tmp/cache", "tmp/sockets", "public/system"

# Default value for default_env is {}
# set :default_env, { path: "/opt/ruby/bin:$PATH" }

# Default value for keep_releases is 5
# set :keep_releases, 5

namespace :deploy do
  after "deploy:install", "nginx:install"
  after "deploy:published", "deploy:setup"
  after "deploy:setup", "nginx:setup"
  after "deploy:finished", "mongo:restart"
  after "mongo:restart", "nginx:restart"
end
