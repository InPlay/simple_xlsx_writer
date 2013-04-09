module SimpleXlsx
  module AttrsSerializer

    def serialize_attrs attrs
      attrs.map{|k,v| (v != nil) ? "#{k.to_s.to_xs}=\"#{v.to_s.to_xs}\"" : nil}.compact.join ' '
    end

  end
end
