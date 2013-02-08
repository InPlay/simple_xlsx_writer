require File.join(File.dirname(__FILE__), '..', 'test_helper')
require "rexml/document"
require 'time'

module SimpleXlsx

class SharedStringsTest < Test::Unit::TestCase

  def test_add_string
    str = ''
    io = StringIO.new str
    doc = Document.new io

    ss = SharedStrings.new(io)

    v = (ss << 'first test string')
    assert_equal v, 0
    v = (ss << 'second test string')
    assert_equal v, 1
    v = (ss << 'first test string')
    assert_equal v, 0

    # it's just the fragment
    doc = REXML::Document.new "<sst>#{str}</sst>"
    si = doc.root.elements
    assert_equal 2, si.to_a.size
    assert_equal doc.root.elements['si'].elements['t'].to_a.first, 'first test string'
  end

end

end
