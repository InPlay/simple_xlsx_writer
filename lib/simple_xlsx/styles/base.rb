module SimpleXlsx
  class Styles
    class Base
      include AttrsSerializer

      def initialize
        @content = {}
      end

      def length
        @content.length
      end

      def << hash
        h = apply_defaults hash
        validate h
        @content[h] ||= @content.length
      end

      private

      def apply_defaults hash
        hash
      end

      def validate hash
      end

      def content
        @content.sort_by{|k,v| v}.map{|h| h[0]}
      end

      def self.validate_color color
        raise ArgumentError, "Invalid color value: \"#{color}\"." unless color.match(/[0-9A-Fa-f]{8}/)
      end

      def self.format_color color
        color.upcase.to_xs
      end

    end
  end
end
