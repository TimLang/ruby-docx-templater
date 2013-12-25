# -*- encoding : utf-8 -*-
require 'rubygems'
require 'nokogiri'
require 'zipruby'
require 'mini_magick'
require 'base64'

module DocxTemplater
  def log(str)
    # braindead logging
    # puts str
  end
  extend self
end

require 'docx_templater/ext/array'
require 'docx_templater/entity/table'
require 'docx_templater/element'
require 'docx_templater/template_processor'
require 'docx_templater/docx_creator'
require 'docx_templater/entity/image'
