require File.join(File.dirname(__FILE__), '..', 'test_helper')

module SimpleXlsx

class AttrSerializerTest < Test::Unit::TestCase

  class TestClass
    include AttrsSerializer
  end

  def test_add_sheet
    t = TestClass.new
    str = t.serialize_attrs :foo => 123,
      :bar => 234,
      :baz => nil

    assert str.include?('foo')
    assert str.include?('123')
    assert str.include?('bar')
    assert str.include?('234')
    assert !str.include?('baz')

    str = t.serialize_attrs :foo => :bar
    assert str == 'foo="bar"'

    str = t.serialize_attrs :foo => true
    assert str == 'foo="true"'

    str = t.serialize_attrs :foo => false
    assert str == 'foo="false"'
  end

end

end
