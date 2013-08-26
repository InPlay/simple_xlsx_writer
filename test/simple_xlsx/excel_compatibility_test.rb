require File.join(File.dirname(__FILE__), '..', 'test_helper')

module SimpleXlsx

class TruncateUriTest < Test::Unit::TestCase

  def test_truncation
    ["%E2%9C%93", "%e2%9c%93"].each {|char|
      (0..9).each{|offset|

        test_uri = "http://google.com?" + ("a" * offset) + (char * 1024)
        t = SimpleXlsx::ExcelCompatibility::truncate_uri(test_uri)

        assert t.length < 256
        assert t.length >= 244
        assert t.end_with?(char), 'last char must be cut completely'

      }
    }
  end

end
end

