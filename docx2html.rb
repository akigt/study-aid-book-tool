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
        @subject_dir = fn.split("/")[-2]
        @xml_dest_dir = @xml_dir + @subject_dir
        @html_dest_dir = @html_dir + @subject_dir
        FileUtils.mkdir_p @xml_dest_dir
        FileUtils.mkdir_p @html_dest_dir
        FileUtils.cp(Dir.glob("docx/#{@subject_dir}/*.png"), @html_dest_dir)
        @xml_name = File.join(@xml_dest_dir, @base_fn.split(".").first + ".xml")
        @html_name = File.join(@html_dest_dir, @base_fn.split(".").first + ".html")
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
        data.gsub!(/#{id}.*?\.png/) { |v|
            %Q[<div class="global--image_container"><img src="#{v}" alt=""></div>]
        }
        # バルーンとセットの画像にタグ付け
        data.gsub!(/([^>]*\/share\/assets\/img\/.*\.png)\n/) {         
            %Q[<dt><img src="#{$1}" alt=""></dt><dd>]
        }
    end

    def removeUnnecessaryTag data
        #不必要に連続するspanタグを消す。classが違い、役割も違うspanタグも消してしまうのでもっと正確な正規表現にしたい
        # data.gsub!(/<\/span><span class=".*?">/,"")
        # data.gsub!(/<dd><br>/,"<dd>")
        # data.gsub!(/%中%\p{blank}?(.*?)/,"テスト　#{$1}")
        data.gsub!(/<span class="(.+?)">([^<>]+?)(?=<\/span><span class="\1">)/){
            "<span class=\"#{$1}\">#{$2}<ToBeDeleted>"}
        data.gsub!(/<ToBeDeleted><\/span><span class="(.+?)">/,"")
        # data.gsub!(/<ToBeDeleted>(.+?)<\/span><span class="(.+?)"><\/ToBeDeleted>/){
        #     "#{$1}"
        # }
        # data.gsub!(/<\/ToBeDeleted><span class=".*?">/,"")
        # while data.sub!(/<span class="(.+?)">(.+?)<\/span>(?=<span class="\1">)/,"<span class=\"#{$1}\">#{$1}#{$2}</ToBeDeleted>") do
        # end
        # data.gsub!(/<\/aside><aside class=".*?">/,"")
    end

    def t_html data
         # 参考書用テキスト、t-htmlのスニペットのタグ
        %Q[<!doctype html>
        <html lang="ja">

        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1, maximum-scale=1">
            <link rel="stylesheet" href="../../../share/assets/css/style.css">
            <link rel="stylesheet" href="../../../share/assets/css/overwrite.css"> </head>

        <body id="index">
            <div class="global--wrapper">
                #{data}
            </div>
            <script src="../../../share/assets/js/jquery-2.1.4.min.js"></script>
            <script src="https://cdn.nnn.ed.nico/MathJax/MathJax.js?config=TeX-MML-AM_CHTML" type="text/javascript"></script>
        </body>

        </html>]
    end

    def t_html_lecture data
         # 授業用テキスト、t-html-lectureのスニペットのタグ
        %Q[<!doctype html>
        <html lang="ja">
        
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, user-scalable=no, initial-scale=1, maximum-scale=1">
            <link rel="stylesheet" href="../../../../share/assets/css/style.css">
            <link rel="stylesheet" href="../../../../share/assets/css/overwrite.css"> </head>
        
        <body id="index">
            <div class="global--wrapper">
              #{data}
            </div>
            <script src="../../../../share/assets/js/jquery-2.1.4.min.js"></script>
            <script src="https://cdn.nnn.ed.nico/MathJax/MathJax.js?config=TeX-MML-AM_CHTML" type="text/javascript"></script>
        </body>
        
        </html>]

    end

    # XMLから抽出した値によって適用するスタイルを変更する
    def style val
        # puts val
        case val

        #見出し類
        when "a0","ab" then #見出し1 
            result = "global--headline_1"
        when "a" then #見出し2
            result = "global--headline_2"
        when "a3" then #見出し3
            result = "global--headline_3"
   
        #メッセージ枠類
        when "af3" then #赤枠
            result = "global--block-message_strong_red"
        when "af7" then #緑枠
            result = "global--block-message_strong_green"
        when "afd" then #青枠 
            result = "global--block-message_strong_blue"
        when "aff" then #灰枠
            result = "global--block-message_strong_gray"
        when "af5" then #黄枠
            result = "global--block-message_strong_yellow"
        when "afb" then #紫枠
            result = "global--block-message_strong_purple"
        when "afff9" then #ふきだし・バルーン用
            result = "global--balloon"
        
        #色付きメッセージブロック
        when "affff7","affffe" then  #黄色のブロック、罫線なし
            result = "global--block-message_yellow"
        when "afffff8" then  #灰色のブロック、罫線なし
            result = "global--block-message_gray"
        
        #ラベル類
        when "afff2" then #赤ラベル 
            result = "global--icon-point_red"
        when "afffffd","afffffc","af6","00FF00" then  #緑ラベル
            result = "global--icon-point_green"
        when "affffff1" then #青ラベル 
            result = "global--icon-point_blue"
        when "affffff2","affffff3" then #紫ラベル 
            result = "global--icon-point_purple"
        when "affffff5" then #灰ラベル 
            result = "global--icon-point_gray"
        when "afffffa" then #黄色ラベル
            result = "global--icon-point_yellow"
            
        #装飾文字類
        when "aff3","aff4","af4","FF0000" then  #赤字
            result = "global--text-red"
        when "affff0","affff1","af8","afe","0000FF","0070C0" then  #青字
            result = "global--text-blue"
        when "affa","affffffb","affffffa" then  #太字 
            result = "global--text-strong"
        when "aff9" then  #大文字 
            result = "global--text-big"
        when "aff2" then  #小文字 
            result = "global--text-small"

        # when "affffff0" then  #公式いろいろ
        #     result = "テスト"
        else
            result = "undefined"
        end
        result
    end

    def surrounding(cssName,inside)
        case cssName
        when /global--headline_([\w]*)/
            "<h#{$1} class=\"global--headline_#{$1}\">" + inside + "</h#{$1}>"
        when /global--icon-point_([\w]*)/
            "<span class=\"global--icon-point_#{$1}\">" + inside + "</span>"
        when /global--text-([\w]*)/
            "<span class=\"global--text-#{$1}\">" + inside + "</span>"
        when /global--block-message_([\w]*)/
            "<aside class=\"global--block-message_#{$1}\">" + inside + "</aside>"
        when "global--balloon"
            "<dl class=\"global--balloon\">" + inside + "</dd></dl>"
        else 
            inside
        end
        
    end

    def parseParagraph elm
        # パラグラフ全体にスタイルがあった場合の処理
        if pStyle = elm.elements[".//w:pStyle"] then 
            pStyleVal = pStyle.attributes["w:val"]
            cssName = style(pStyleVal)
            surrounding(cssName,parseStyle(elm))
        # スタイルのない通常のパラグラフのための処理
        else
            "<p>" + parseStyle(elm) + "</p>"
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
                    #.//w:tでは最初のtのみを取得するので二つ以上あると一部の文字が消える
                    "\n" + e.get_elements(".//w:t")[1].text
                else
                    normalText.text
                end
            end
            }.join("").chomp("")
    end

    #表の行をパースするための関数
    def parseTableRow elm
        #テーブル内の各行要素を取り出し<tr>タグで囲んで、列をパースする処理を行う
         elm.get_elements(".//w:tr").to_a.map.with_index { |row, i|
            "<tr>" + parseTableColumn(row,i) + "</tr>"
            }.join("").chomp("")
    end

    #表の列をパースするための関数。第二引数で行のインデックスiを受け取りthタグとtdタグを使い分ける
    def parseTableColumn(elm,i)
        #テーブルの行内部の列の要素を取り出し<td>タグで囲んで、内部のテキストをパースする処理を行う
         elm.get_elements(".//w:tc").to_a.map { |column|
            cellData = parseParagraph(column)
            #装飾タグがないセルにはpタグを付与する
            if !cellData.include?("</") and !(cellData == "") then
                cellData = "<p>" + cellData + "</p>"
            end
            #最初の行i=0だったらthタグを使用する。今のところthタグの必要性が不明なため-1に。
            if i == -1 then
                "<th>" + cellData + "</th>"
            else
                "<td>" + cellData + "</td>"
            end
            }.join("").chomp("")
    end

    def parse
        # OpenXMLをパースしてテキストを抽出
        doc = REXML::Document.new(check_br(@xml_name))

        data = REXML::XPath.match(doc.root,'//*[self::w:tbl or self::w:p[not(ancestor::w:tbl)]]').map { |elm|
            
            # テーブルがあったときの処理
            if elm.name == "tbl" then
                "<table>" + parseTableRow(elm) + "</table>"
            # パラグラフを処理
            else
                parseParagraph(elm)
            end

        }.join("").chomp("")

        # removeUnnecessaryTag(data)
        add_img(data)
        data.gsub!(/\n/,"<br>\n")
        # add_img(data)
        removeUnnecessaryTag(data)

        #与えられたコマンドライン引数によって使用する雛形を選択
        case ARGV[0]
        when "lecture"
            # t-html-lectureのスニペットのタグ
            t_html_lecture(data)
        else
            # t-htmlのスニペットのタグ
            t_html(data)
        end


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
    data = Dir.glob("docx/*/*.docx")
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
