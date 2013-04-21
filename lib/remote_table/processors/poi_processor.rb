class RemoteTable
  module PoiProcessor
    TAG = /<[^>]+>/
    BLANK = ''

    def _each
      book = processor_class.open local_copy.path
      sheet = book.worksheets[0]
      @rows = sheet.rows

      (first_row..last_row).each do |i|
        result = headers ? row_in_hash(@rows[i]) : row_in_array(@rows[i])

        yield result if result
      end
    ensure
      local_copy.cleanup
    end

    private
    def row_in_array(row)
      some_value_present = false
      output = (0..(row.last_cell_num)).map do |i|
        memo = row[i].to_s.dup
        memo = assume_utf8 memo
        memo.gsub! TAG, BLANK
        memo.strip!
        if not some_value_present and not keep_blank_rows and memo.present?
          some_value_present = true
        end
        memo
      end
      if keep_blank_rows or some_value_present
        output
      end
    end

    def row_in_hash(row)
      some_value_present = false
      output = ::ActiveSupport::OrderedHash.new
      current_headers.each do |k, x|
        memo = row[x].to_s.dup
        memo = assume_utf8 memo
        memo.gsub! TAG, BLANK
        memo.strip!
        if not some_value_present and not keep_blank_rows and memo.present?
          some_value_present = true
        end
        output[k] = memo
      end
      if keep_blank_rows or some_value_present
        output
      end
    end

    def first_row
      return @first_row if @first_row

      @first_row = crop ? crop.first : skip
      headers ? @first_row += 1 : @first_row
    end

    def last_row
      @last_row ||= crop ? crop.last - 1 : @rows.count - 1
    end

    def current_headers
      return @current_headers if @current_headers

      @current_headers = ::ActiveSupport::OrderedHash.new
      if headers == :first_row
        header_index = first_row - 1
        (0..(@rows[header_index].last_cell_num - 1)).each do |x|
          v = @rows[header_index][x]
          if v.present?
            v = assume_utf8 v
            # 'foobar' is found at column 6
            @current_headers[v.to_s] = x
          end
        end
        # "advance the cursor"
        @first_row += 1
      else
        headers.each_with_index do |k, i|
          @current_headers[k] = i + 1
        end
      end

      @current_headers
    end
  end
end
