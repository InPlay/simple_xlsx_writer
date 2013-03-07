module SimpleXlsx
  module AttrsSerializer

    def serialize_attrs attrs
      attrs.map{|k,v| "#{k.to_s.to_xs}=\"#{v.to_s.to_xs}\""}.compact.join ' '
    end

  end
end
