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
        # バルーンとセットの画像にタグ付け
        data.gsub!(/^.*\/share\/assets\/img\/.*\.png/) { |v|        
            %Q[<dt><img src="#{v}" alt=""></dt><dd>]
        }
    end

    def removeUnnecessaryTag data
        data.gsub!(/<\/span><span class=".*?">/,"")
        # data.gsub!(/<\/aside><aside class=".*?">/,"")
    end

    # XMLから抽出した値によって適用するスタイルを変更する
    #<w:rStyle w:val="af8"/> 8は青 6は緑？ 4は赤
    def style val
        # puts val
        case val
        #2A1ACC1F paraID波線 pstyle aff
        when "a0","ab" then #見出し1 
            result = "global--headline_1"
        when "a" then #見出し2
            result = "global--headline_2"
        # when "af1","af2" then #アノテーション
        #     result = "annotation"
        when "af3" then #赤枠
            result = "global--block-message_strong_red"
        when "afd" then #青枠
            result = "global--block-message_strong_blue"
        when "aff" then #灰枠
            result = "global--block-message_strong_gray"
        when "afff9" then #波線の枠
            result = "global--balloon"
        #ラベル類
        when "afff2" then #赤ラベル 
            result = "global--icon-point_red"
        when "afffffd","afffffc","af5","af6","00FF00" then  #緑ラベル
            result = "global--icon-point_green"
        when "affffff1" then #青ラベル 
            result = "global--icon-point_blue"
        when "affffff3" then #紫ラベル 
            result = "global--icon-point_purple"
        when "affffff5" then #灰ラベル 
            result = "global--icon-point_gray"
        when "aff4","af4","FF0000" then  #赤字
            result = "global--text-red"
        when "affff1","af7","af8","afe","0000FF","0070C0" then  #青字
            result = "global--text-blue"
        # when "af9" then  #太字 
        #     result = "strong"
        # when "aff" then  #斜体 
        #     result = "em"
        when "affffff0" then  #公式いろいろ
            result = "テスト"
        when "affff7","affffe" then  #黄色の枠 
            result = "global--block-message_yellow"
        else
            result = "undefined"
        end
        result
    end

    def surrounding(cssName,inside)
        case cssName
        when /global--headline_([\w]*)/
            "<h#{$1} class=\"global--headline_#{$1}\">" + inside + "</h#{$1}>\n"
        when /global--icon-point_([\w]*)/
            "<span class=\"global--icon-point_#{$1}\">" + inside + "</span>"
        when /global--text-([\w]*)/
            "<span class=\"global--text-#{$1}\">" + inside + "</span>"
        when "テスト"
            "<AAAAA>" + inside + "</AAAAA>\n"
        when /global--block-message_([\w]*)/
            "<aside class=\"global--block-message_#{$1}\">" + inside + "</aside>"
        when "global--balloon"
            "<dl class=\"global--balloon\">\n" + inside + "</dd></dl>"
            # "<AAAAA>" + inside + "</AAAAA>\n"
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
            }.join("").chomp("") + "\n"
    end

    def parse
        # OpenXMLをパースしてテキストだけ抽出
        doc = REXML::Document.new(check_br(@xml_name))

        data = doc.elements.to_a("//w:p").map { |elm|

            #　パラグラフ全体にスタイルがあった場合の処理
            if rStyle = elm.elements[".//w:pStyle"] then 
                rStyleVal = rStyle.attributes["w:val"]
                cssName = style(rStyleVal)
                surrounding(cssName,parseStyle(elm))
            # 通常のパラグラフのための処理
            else
                parseStyle(elm) + "\n"
            end

        }.join("").chomp("")

        removeUnnecessaryTag(data)
        data.gsub!(/\n/,"<br>\n")
        add_img(data)

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
