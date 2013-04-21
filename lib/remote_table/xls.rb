class RemoteTable
  # Parses XLS files using Roo's Excel class.
  module Xls
    def self.extended(base)
      #base.extend ProcessedByRoo
      base.extend RUBY_PLATFORM == 'java' ? PoiProcessor : ProcessedByRoo
    end

    def roo_class
      Roo.const_defined?(:Excel) ? Roo::Excel : ::Excel
    end

    def processor_class
      POI::Workbook
    end
  end
end
