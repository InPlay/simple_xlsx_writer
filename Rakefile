require 'rake'
require 'rake/testtask'

task :default => [:test]

Rake::TestTask.new do |test|
  test.libs       << "test"
  test.test_files =  Dir['test/**/*_test.rb'].sort
  test.verbose    =  true
end

desc "generate tags for emacs"
task :tags do
  sh "ctags -Re lib/ "
end

task :build do
  system "gem build simple_xlsx_writer.gemspec"
end
