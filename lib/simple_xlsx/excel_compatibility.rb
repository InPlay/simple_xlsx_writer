module SimpleXlsx
  module ExcelCompatibility

    UTF_CONT = /\%[8-9abAB]/
    UTF_FIRST = /\%[cdefCDEF]/

    def self.truncate_uri uri
      r = uri[0..254]
      return r unless r.length == 255

      # cutting invalid parts
      ri = r.rindex('%')
      r = r[0..ri-1] if ri && ri >= (r.length-2)

      # cutting all utf-low
      ri = r.rindex UTF_CONT

      while ri && ri >= (r.length-3)
        r = r[0..ri-1]
        ri = r.rindex UTF_CONT
      end

      ri = r.rindex UTF_FIRST
      r = r[0..ri-1] if ri && ri >= (r.length-3)

      r
    end

  end
end
