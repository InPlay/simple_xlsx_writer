module SimpleXlsx
  class Styles
    class CellStyles < Base

      def to_stream stream
        stream.puts "<cellStyles count=\"#{length}\">"
        content.each{|(c,id)| 
          stream.puts "<cellStyle builtinId=\"#{c[:builtin_id].to_s.to_xs}\" customBuiltin=\"#{c[:custom_builtin].to_s.to_xs}\" name=\"#{c[:name].to_xs}\" xfId=\"#{c[:xf_id].to_s.to_xs}\"/>"
        }
        stream.puts "</cellStyles>"
      end


      def validate o
        super
        raise ArgumentError, "No builtin_id specified" unless o.has_key? :builtin_id
        raise ArgumentError, "No custom_builtin specified" unless o.has_key? :custom_builtin
        raise ArgumentError, "No xf_id specified" unless o.has_key? :xf_id
      end

    end
  end
end


