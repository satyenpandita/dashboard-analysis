# require 'capistrano-virtualenv'

server "ec2-52-43-105-252.us-west-2.compute.amazonaws.com", :web, :app, :db,primary: true

set :tld, "ec2-52-43-105-252.us-west-2.compute.amazonaws.com"

# Uncomment the packages that you require. Do not comment out any package
load "config/recipes/base"

# RVM settings
load "config/recipes/rvm_signature_fix"
# load "config/recipes/rvm"

# Nginx
load "config/recipes/nginx"
# load "config/recipes/nginx_with_https"

# Unicorn
load "config/recipes/unicorn"

#Flask
load "config/recipes/flask"


set :application, "auroville"
set :repository,  "git@github.com:Prashant31/dashboard-analysis.git"

set :user, "ubuntu"
set :deploy_to, "/home/#{user}/apps/#{application}"
set :deploy_via, :remote_cache
set :use_sudo, false
set :scm, "git"
default_run_options[:pty] = true
set :ssh_options, { forward_agent: true, paranoid: true, keys: "~/.ssh/id_rsa" }