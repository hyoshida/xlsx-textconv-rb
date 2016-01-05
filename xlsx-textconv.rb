#!/usr/bin/env ruby
require 'roo'

class XlsxTextconv
  class << self
    def run(filepath)
      open(filepath).each_with_pagename {|name, sheet| puts name; puts sheet.to_yaml}
    end

    private

    def open(filepath)
      klass = class_for(filepath)
      klass.new(filepath)
    end

    def class_for(fillepath)
      case File.extname(fillepath)
      when '.xls'
        require 'roo-xls'
        Roo::Excel
      when '.xlsx'
        Roo::Excelx
      else
        fail "Can't open the type of file: #{fillepath}"
      end
    end
  end
end

XlsxTextconv.run(ARGV[0])
