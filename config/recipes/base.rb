def template(from, to)
  erb = File.read(File.expand_path("../templates/#{from}", __FILE__))
  put ERB.new(erb).result(binding), to
end

def set_default(name, *args, &block)
  set(name, *args, &block) unless exists?(name)
end

def remote_file_exists?(full_path)
  'true' == capture("if [ -e #{full_path} ]; then echo 'true'; fi").strip
end

namespace :deploy do
  desc "Install everything onto the server"
  task :install do
    run "sudo apt-get -y update"
    run "sudo apt-get -y upgrade"
    run "sudo apt-get -y install  python python3-pip python-virtualenv nginx gunicorn"
  end
end
