# -*- encoding: utf-8 -*-

lib = File.expand_path('../lib/', __FILE__)
$:.unshift lib unless $:.include?(lib)

require 'rubygems'

Gem::Specification.new do |s|
  s.name = "simple_xlsx_writer"
  s.version = "0.5.3.undev"
  s.authors = ["Dee Zsombor", "Victor Bilyk"]
  s.email = ["zsombor@primalgrasp.com", "victorbilyk@gmail.com"]
  s.homepage = "http://simplxlsxwriter.rubyforge.org"
  s.rubyforge_project = "simple_xlsx_writer"
  s.platform = Gem::Platform::RUBY
  s.summary = "Just as the name says, simple writter for Office 2007+ Excel files"
  s.files = Dir["{bin,lib}/**/*"] + Dir["LICENSE"]+ Dir["Rakefile"]
  s.require_path = "lib"
  s.test_files = Dir["{test}/**/*test.rb"] + Dir["test/test_helper.rb"]
  s.has_rdoc = true
  s.extra_rdoc_files = Dir["README"]
  s.add_dependency("rubyzip", ">= 1.0.0")
  s.add_dependency("fast_xs", ">= 0.7.3")
end
