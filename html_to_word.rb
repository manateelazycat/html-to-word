#coding: utf-8
require 'base64'
require 'cgi'
require 'digest/sha1'
require 'fastimage'

module HTMLToWord
  PAGE_VIEW_HTML_TEMPLATE = "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\""

  PAGE_VIEW_HEAD_TEMPLATE = "<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:TrackMoves>false</w:TrackMoves><w:TrackFormatting/><w:ValidateAgainstSchemas/><w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid><w:IgnoreMixedContent>false</w:IgnoreMixedContent><w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText><w:DoNotPromoteQF/><w:LidThemeOther>EN-US</w:LidThemeOther><w:LidThemeAsian>ZH-CN</w:LidThemeAsian><w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript><w:Compatibility><w:BreakWrappedTables/><w:SnapToGridInCell/><w:WrapTextWithPunct/><w:UseAsianBreakRules/><w:DontGrowAutofit/><w:SplitPgBreakAndParaMark/><w:DontVertAlignCellWithSp/><w:DontBreakConstrainedForcedTables/><w:DontVertAlignInTxbx/><w:Word11KerningPairs/><w:CachedColBalance/><w:UseFELayout/></w:Compatibility><w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel><m:mathPr><m:mathFont m:val=\"Cambria Math\"/><m:brkBin m:val=\"before\"/><m:brkBinSub m:val=\"--\"/><m:smallFrac m:val=\"off\"/><m:dispDef/><m:lMargin m:val=\"0\"/> <m:rMargin m:val=\"0\"/><m:defJc m:val=\"centerGroup\"/><m:wrapIndent m:val=\"1440\"/><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/></m:mathPr></w:WordDocument></xml><![endif]-->\n"

  def self.convert(html,
                   document_guid,
                   max_image_width,
                   css_files,
                   image_filter,
                   percent_start_number,
                   percent_end_number,
                   percent_update,
                   development_proxy_address,
                   development_proxy_port,
                   production_proxy_address,
                   production_proxy_port)
    # Make sure all images' size small than max size.
    no_size_attr_hash = Hash.new

    doc = Nokogiri::HTML(html)
    doc.css("img").each do |img|
      if (img.keys.include? "width") && (img.keys.include? "height")
        img_width = img["width"].to_i
        img_height = img["height"].to_i

        # Image won't show in Word file if image's width or height equal 0.
        if (img_width != 0) && (img_height != 0)
          render_width = [img_width, max_image_width].min
          render_height = render_width * 1.0 / img_width * img_height

          img["width"] = render_width
          img["height"] = render_height
        end
      else
        no_size_attr_hash[img["src"]] = nil
      end
    end

    # Unescape html first to avoid base64's link not same as image tag's link.
    html = CGI.unescapeHTML(doc.to_html)

    download_image_count = doc.css("img").length
    if download_image_count > 0
      percent_update.call(percent_start_number)
    end

    # Fetch image's base64.
    base64_cache = Hash.new
    filter_image_replace_hash = Hash.new
    mhtml_bottom = "\n"
    download_image_index = 0
    Nokogiri::HTML(html).css('img').each do |img|
      if img.keys.include? "src"
        # Init.
        image_src = img.attr("src")
        begin
          uri = URI(image_src)
          proxy_addr = nil
          proxy_port = nil

          # Use image_filter to convert internal images to real image uri.
          real_image_src = image_filter.call(image_src)
          base64_image_src = image_src
          if real_image_src != image_src
            uri = URI(real_image_src)

            # We need use convert image to hash string make sure all internal image visible in Word file.
            uri_hash = Digest::SHA1.hexdigest(image_src)
            placeholder_uri = "https://placeholder/#{uri_hash}"
            filter_image_replace_hash[image_src] = placeholder_uri
            base64_image_src = placeholder_uri
          else
            # Use proxy when image is not store inside of webside.
            # Of course, you don't need proxy if your code not running in China.
            if %w(test development).include?(Rails.env.to_s)
              proxy_addr = development_proxy_address
              proxy_port = development_proxy_port
            else
              proxy_addr = production_proxy_address
              proxy_port = production_proxy_port
            end
          end

          # Fetch image's base64.
          image_base64 = ""

          if base64_cache.include? image_src
            # Read from cache if image has fetched.
            image_base64 = base64_cache[image_src]
          else
            # Get image response.
            # 
            # URI is invalid if method request_uri not exists.
            if uri.respond_to? :request_uri
              response = Net::HTTP.start(uri.hostname, uri.port, proxy_addr, proxy_port, use_ssl: uri.scheme == "https") do |http|
                http.request(Net::HTTP::Get.new(uri.request_uri))
              end

              if response.is_a? Net::HTTPSuccess
                image_base64 = Base64.encode64(response.body)

                base64_cache[image_src] = image_base64
              end
            end
          end

          # Fetch image size if img tag haven't any size attributes.
          if (no_size_attr_hash.include? image_src) && (no_size_attr_hash[image_src] == nil)
            proxy_for_fast_image = nil
            if proxy_addr && proxy_port
              proxy_for_fast_image = "http://#{proxy_addr}:#{proxy_port}"
            end

            # NOTE:
            # This value maybe nil if remote image unreachable.
            no_size_attr_hash[image_src] = FastImage.size(real_image_src, { proxy: proxy_for_fast_image })
          end

          # Build image base64 template.
          if image_base64 != ""
            mhtml_bottom += "--NEXT.ITEM-BOUNDARY\n"
            mhtml_bottom += "Content-Location: #{base64_image_src}\n"
            mhtml_bottom += "Content-Type: image/png\n"
            mhtml_bottom += "Content-Transfer-Encoding: base64\n\n"
            mhtml_bottom += "#{image_base64}\n\n"
          else
            print("Can't fetch image base64: " + base64_image_src + "\n")
          end
        rescue URI::InvalidURIError, URI::InvalidComponentError
          Rails.logger.info "[FILE] Document #{document_guid} contain invalid url, pass it. error: #{e}; backtraces:\n #{e.backtrace.join("\n")}"
        end

        # Update download index to calcuate percent.
        download_image_index += 1
        percent_update.call(percent_start_number + (percent_end_number - percent_start_number) * (download_image_index * 1.0 / download_image_count))
      end
    end
    mhtml_bottom += "--NEXT.ITEM-BOUNDARY--"

    # Adjust image size of img tag that haven't size attributes.
    doc = Nokogiri::HTML(html)
    doc.css("img").each do |img|
      if img.keys.include? "src"
        # no_size_attr_hash[img["src"]] will got nil if remote image unreachable.
        # So give up scale image size here because the image won't show up in Word.
        if no_size_attr_hash[img["src"]].present?
          size = no_size_attr_hash[img["src"]]

          render_width = [size.first, max_image_width].min
          render_height = render_width * 1.0 / size.first * size.second

          img["width"] = render_width
          img["height"] = render_height
        end
      end
    end
    html = CGI.unescapeHTML(doc.to_html)

    # Replace image hash.
    filter_image_replace_hash.each do |key, value|
      html = html.gsub key, value
    end

    # Pick up style content from stylesheet file.
    stylesheet = ""
    css_files.each do |scss_file|
      if %w(test development).include?(Rails.env.to_s)
        stylesheet += Rails.application.assets.find_asset(scss_file).source
      else
        stylesheet += File.read(File.join(Rails.root, "public", ActionController::Base.helpers.asset_url(scss_file.ext("css"))))
      end
    end

    if download_image_count > 0
      percent_update.call(percent_end_number)
    end

    # Return word content.
    head = "<head>\n #{PAGE_VIEW_HEAD_TEMPLATE} <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\n"
    head += "<style>\n #{stylesheet} _\n</style>\n</head>\n"
    body = "<body> #{html} </body>"

    mhtml_top = "Mime-Version: 1.0\nContent-Base: #{document_guid} \n"
    mhtml_top += "Content-Type: Multipart/related; boundary=\"NEXT.ITEM-BOUNDARY\";type=\"text/html\"\n\n--NEXT.ITEM-BOUNDARY\n"
    mhtml_top += "Content-Type: text/html; charset=\"utf-8\"\nContent-Location: #{document_guid} \n\n"
    mhtml_top += "<!DOCTYPE html>\n<html #{PAGE_VIEW_HTML_TEMPLATE}  >\n #{head} #{body} </html>"

    mhtml_top + mhtml_bottom
  end
end
