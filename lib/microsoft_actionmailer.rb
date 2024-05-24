require "microsoft_actionmailer/version"
require 'microsoft_actionmailer/railtie' if defined?(Rails)
require 'microsoft_actionmailer/api'

require 'httparty'
require 'net/http'
require 'uri'

module MicrosoftActionmailer

  GRAPH_HOST = 'https://graph.microsoft.com'.freeze

  class DeliveryMethod
    include MicrosoftActionmailer::Api

    attr_reader :access_token
    attr_reader :delivery_options
    attr_reader :api_user_dir

    def initialize params
      begin
        @access_token = params[:authorization]
        @delivery_options = params[:delivery_options] || {}
        @api_user_dir = params[:user_id].blank? ? "me" : "users/#{params[:user_id]}"
      rescue => e
        raise "MicrosoftActionmailer error in initialize: #{e.message}"
      end
    end

    def deliver! mail
      begin
        if mail.html_part.present?
          body = mail.html_part.body.encoded
        else
          body = mail.body.encoded
        end
  
        message = ms_create_message(
          access_token,
          mail.subject,
          body,
          mail.to,
          mail.cc,
          mail.bcc,
          mail.reply_to,
          mail.attachments,
          api_user_dir
        )
  
        before_send = delivery_options[:before_send]
        if before_send && before_send.respond_to?(:call)
          before_send.call(mail, message)
        end
  
        ms_send_message(access_token, message['id'], api_user_dir)
  
        after_send = delivery_options[:after_send]
        if after_send && after_send.respond_to?(:call)
          after_send.call(mail, message)
        end
      rescue => e
        raise "MicrosoftActionmailer error in deliver: #{e.message}"
      end 
    end
  end
end
