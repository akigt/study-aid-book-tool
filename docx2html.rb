#!/usr/bin/env ruby
# coding: utf-8
require 'rexml/document'
require 'fileutils'
require 'zip' # gem install rubyzip

def main; yield end

class Doc2Txt
    attr_accessor :xml_dir, :txt_dir

    def initialize
        @xml_dir = Dir.pwd.encode("utf-8") + "/xml/"
        @html_dir = Dir.pwd.encode("utf-8") + "/html/"
    end

    def set fn
        @fn = fn
        @base_fn = fn.split("/").last
        FileUtils.mkdir_p @xml_dir
        FileUtils.mkdir_p @html_dir
        @xml_name = File.join(@xml_dir, @base_fn.split(".").first + ".xml")
        @html_name = File.join(@html_dir, @base_fn.split(".").first + ".html")
    end

    def extract_xml
        # docxからdocument.xmlを取り出す
        Zip::File.open(@fn) { |z|
            entry = z.glob("word/document.xml").first
            entry.extract(@xml_name) { true }
        }
    end

    def doc_header
        %Q[<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body>]
    end

    def doc_footer
        %Q[</w:body></w:document>]
    end

    # <w:p> -> 単体で見て後ろに改行
    # <w:br/> -> 改行
    # w:body/w:br -> w:body/w:p [<w:p/>]
    # w:body/w:p//w:br -> [<w:t>\n</w:t>]
    
    def check_br fn
        doc = REXML::Document.new(open(fn))
        str = doc.elements.to_a("w:document/w:body/*").map { |e|
            str = e.to_s
                .gsub(/^<w:br\/>/, "<w:p/>")
                .gsub(/^<w:br>/, "<w:p/>")
                .gsub("<w:br/>", "<w:t>\n</w:t>")
                .gsub("<w:br>", "<w:t>\n</w:t>")
            str
        }.join("")
        str = doc_header + str + doc_footer
        str
    end

    def add_img data
        id = @base_fn.split(".").first
        #通常の画像にimgタグを付ける
        data.gsub!(/#{id}.*\.png/) { |v|
            %Q[<div class="global--image_container">
            <img src="#{v}" alt="">
            </div>
            ]
        }
        #バルーンとセットの画像にタグ付け
        data.gsub!(/^.*\/share\/assets\/img\/.*\.png/) { |v|        
            %Q[<dl class="global--balloon"><dt><img src="#{v}" alt=""></dt>
              <dd>
                セリフ
              </dd>
            </dl>
            ]
        }
    end

    def add_tex data
        data.gsub!(/<\/span><span class=".*?">/,"")
    end

    # XMLから抽出した値によって適用するスタイルを変更する
    #<w:rStyle w:val="af8"/> 8は青 6は緑？ 4は赤
    def style val
        puts val
        case val
        when "a0","ab" then #見出し1 
            result = "global--headline_1"
        when "a" then #見出し2
            result = "global--headline_2"
        when "af1","af2" then #アノテーション
            str = "annotation"
        when "aff4","af4","FF0000" then  #赤
            result = "global--text-red"
        when "afffffc","af5","af6","00FF00" then  #例
            result = "global--icon-point_green"
        when "affff1","af7","af8","afe","0000FF","0070C0" then  #青
            result = "global--text-blue"
        when "af9" then  #太字 
            str = "strong"
        when "aff" then  #斜体 
            str = "em"
        when "affffff0" then  #公式いろいろ
            result = "テスト"
        when "affff7" then  #黄色の枠 
            result = "global--block-message_yellow"
        else
            result = "aa"
        end
        result
    end

    def surrounding(cssName,inside)
        case cssName
        when /global--headline_([\w]*)/
            "<h#{$1} class=\"global--headline_#{$1}\">" + inside + "</h#{$1}>\n"
        when /global--icon-point_([\w]*)/
            "<span class=\"global--icon-point_#{$1}\">" + inside + "</span>\n"
        when /global--text-([\w]*)/
            "<span class=\"global--text-#{$1}\">" + inside + "</span>"
        when "テスト"
            "<AAAAA>" + inside + "</AAAAA>\n"
        when /global--block-message_([\w]*)/
            "<aside class=\"global--block-message_#{$1}\">" + inside + "</aside>"
        else 
            inside
        end
        
    end

    def parseStyle elm
        #XMLの要素からスタイルを抜き取りテキストにspanタグをくっつける
         elm.get_elements(".//w:r").to_a.map { |e| 
            #教科によって文字の色のつけ方が違うので場合分け・・・
            if (rStyle = e.elements[".//w:color"]) or (rStyle = e.elements[".//w:rStyle"]) then 
                rStyleVal = rStyle.attributes["w:val"]
                cssName = style(rStyleVal)
                surrounding(cssName,e.elements[".//w:t"].text.gsub(/\n/, ""))
            #スタイルの無い通常の文の場合
            elsif normalText = e.elements[".//w:t"] then 
                if e.get_elements(".//w:t")[1] then #一つのw:rタグ内にw:tが二つあるケースに対処するため。
                    #.//w:tでは最初のtのみを取得するので一部の文字が消える
                    "\n" + e.get_elements(".//w:t")[1].text
                else
                    normalText.text
                end
            end
            }.join("").chomp("")
    end

    def parse
        # OpenXMLをパースしてテキストだけ抽出
        doc = REXML::Document.new(check_br(@xml_name))

        data = doc.elements.to_a("//w:p").map { |elm|

            #　パラグラフ全体にスタイルがあった場合の処理
            if rStyle = elm.elements[".//w:pStyle"] then 
                rStyleVal = rStyle.attributes["w:val"]
                # puts rStyleVal
                cssName = style(rStyleVal)
                # if cssName
                    surrounding(cssName,parseStyle(elm))
                # else
                    # puts !!cssName
                    # elm.get_elements(".//w:t").to_a.map { |e| e.text }.join("").chomp("") + "\n"
                # end

            # 通常のパラグラフのための処理
            else
                parseStyle(elm) + "\n"
            end

        }.join("").chomp("")

        data.gsub!(/\n/,"<br>\n")
        add_img(data)
        add_tex(data)

        # t-htmlのスニペットのタグ
        '   <!doctype html>
        <html lang="ja">

        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1, maximum-scale=1">
            <link rel="stylesheet" href="../../../share/assets/css/style.css">
            <link rel="stylesheet" href="../../../share/assets/css/overwrite.css"> </head>

        <body id="index">
            <div class="global--wrapper">
                %s
            </div>
            <script src="../../../share/assets/js/jquery-2.1.4.min.js"></script>
            <script src="https://cdn.nnn.ed.nico/MathJax/MathJax.js?config=TeX-MML-AM_CHTML" type="text/javascript"></script>
        </body>

        </html>' % data
        
    end

    def output data
        open(@html_name, "w") { |f| f.write data }
    end

    def main
        extract_xml
        output(parse)
    end
end

if __FILE__ == $PROGRAM_NAME
main {
    dt = Doc2Txt.new
    #data = ARGV.select { |fn| fn =~ /.docx$/ }
    data = Dir.glob("docx/*.docx")
    max_num = data.length
    n = 0
    data.each { |fn|
        dt.set fn
        dt.main
        n += 1
        print "\r[#{n}/#{max_num}]"
    }
    puts
}
end
