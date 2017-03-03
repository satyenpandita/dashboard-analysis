namespace :newrelic do
  namespace :server do
    desc "Set new relic server monitoring"
    task :install do
      run "#{sudo} sh -c 'echo deb http://apt.newrelic.com/debian/ newrelic non-free >> /etc/apt/sources.list.d/newrelic.list'"
      run "#{sudo} sh -c 'wget -O- https://download.newrelic.com/548C16BF.gpg | apt-key add -'"
      run "#{sudo} apt-get -y update"
      run "#{sudo} apt-get -y install newrelic-sysmond"
      run "#{sudo} nrsysmond-config --set license_key=c72265339a877f18fd77b0d7a10e9945a80ec7f0"
    end
    after "deploy:install", "newrelic:server:install"

    %w[start stop restart].each do |command|
      desc "#{command} new relic server"
      task command, roles: :web do
        run "#{sudo} /etc/init.d/newrelic-sysmond #{command}"
      end
    end
    after "deploy:restart", "newrelic:server:restart"
  end
end

require 'new_relic/recipes'

# This goes out even if the deploy fails, sadly.
after "deploy",            "newrelic:notice_deployment"
after "deploy:update",     "newrelic:notice_deployment"
after "deploy:migrations", "newrelic:notice_deployment"
after "deploy:cold",       "newrelic:notice_deployment"