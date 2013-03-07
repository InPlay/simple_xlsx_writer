module SimpleXlsx
  class Styles
    module Xf

      extend AttrsSerializer

      ALIGNMENT_GENERAL = :general
      ALIGNMENT_LEFT = :left
      ALIGNMENT_CENTER = :center
      ALIGNMENT_RIGHT = :right
      ALIGNMENT_FILL = :fill
      ALIGNMENT_JUSTIFY = :justify
      ALIGNMENT_CENTER_CONTINUOUS = :centerContinuous
      ALIGNMENT_DISTRIBUTED = :distributed

      ALIGNMENT_TOP = :top
      ALIGNMENT_BOTTOM = :top

      HORIZONTAL_ALIGNMENTS = [
        ALIGNMENT_GENERAL, ALIGNMENT_LEFT,
        ALIGNMENT_RIGHT, ALIGNMENT_CENTER, ALIGNMENT_FILL,
        ALIGNMENT_JUSTIFY, ALIGNMENT_CENTER_CONTINUOUS,
        ALIGNMENT_DISTRIBUTED
      ]

      VERTICAL_ALIGNMENTS = [
        ALIGNMENT_TOP, ALIGNMENT_BOTTOM, ALIGNMENT_CENTER,
        ALIGNMENT_JUSTIFY, ALIGNMENT_DISTRIBUTED
      ]
      

      def self.write stream, c

        xf_attrs = {:applyAlignment=>c[:apply_alignment],
          :applyBorder=>c[:apply_border],
          :applyFont=>c[:apply_font],
          :applyProtection=>c[:apply_protection],
          :borderId=>c[:border_id],
          :fillId=>c[:fill_id],
          :fontId=>c[:font_id],
          :numFmtId=>c[:num_fmt_id],
          :xfId=>c[:xf_id]}

        stream.write "<xf #{serialize_attrs xf_attrs}>"
        if a=c[:alignment]
          alignment_attrs = {
            :wrapText=>a[:wrap_text],
            :horizontal=>a[:horizontal],
            :vertical=>a[:vertical],
          }
          stream.write "<alignment #{serialize_attrs alignment_attrs}/>"
        end
        stream.puts "</xf>"
      end

      private

      def validate o
        if (a = o[:alignment])
          h = a[:horizontal]
          raise ArgumentError, "Invalid horizontal alignment #{h}" unless !h || HORIZONTAL_ALIGNMENTS.include?(h)
          v = a[:vertical]
          raise ArgumentError, "Invalid horizontal alignment #{v}" unless !v || VERTICAL_ALIGNMENTS.include?(v)
        end
      end

    end
  end
end



