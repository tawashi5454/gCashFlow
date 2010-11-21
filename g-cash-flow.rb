//
// gCashFlow ver 0.0.1
// Copyright 2010, Takafumi IWAI <takafumi@tawashi.org>
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are
// met:
//
//     * Redistributions of source code must retain the above copyright
// notice, this list of conditions and the following disclaimer.
//     * Redistributions in binary form must reproduce the above
// copyright notice, this list of conditions and the following disclaimer
// in the documentation and/or other materials provided with the
// distribution.
//     * Neither the name of Google Inc. nor the names of its
// contributors may be used to endorse or promote products derived from
// this software without specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
// "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
// LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
// A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
// OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
// SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
// LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
// DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
// THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
// OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

$KCODE = 'u'

require 'rubygems'
require 'gmail'
require 'mail'
require 'kconv'
require 'date'
require "google_spreadsheet"


# Please change these
USER = "username"
PASS = "password"
URL = "https://spreadsheets.google.com/spreadsheet"


class AccoutItem
  attr_accessor :amount, :place, :detail, :date
  
  def initialize(email)
    body = Kconv.toutf8 email.body.to_s
    properties = body.split(/[,\s、　]/)

    @amount = properties[0].to_i || 0
    @place = properties[1] || "" 
    @detail = properties[2] || "" 
    @date = properties[3] ?
      Date.strptime("2010/" + properties[3], "%Y/%m/%d") : email.date
  end

  def year
    @date.year
  end

  def month
    @date.month
  end

  def day
    @date.day
  end

  def to_s
    str = StringIO.new
    str << "--Account Item--\n"
    str << "amount: #{@amount}\n"
    str << "place: #{@place}\n"
    str << "detail: #{@detail}\n"
    str << "date: #{@date.year}/#{@date.month}/#{@date.day}"
    str.string
  end
end

worksheet_title = Time.now.strftime("%Y/%m")

session = GoogleSpreadsheet.login(USER, PASS)
spreadsheet = session.spreadsheet_by_url(URL)
worksheet = nil

spreadsheet.worksheets.each do |ws|
  next if ws.title != worksheet_title
  worksheet = ws
end 
worksheet ||= spreadsheet.add_worksheet worksheet_title

row_index = 2
while !worksheet[row_index, 1].empty?
  row_index = row_index + 1
end

Gmail.new(USER + '@gmail.com', PASS) do |gmail|
  gmail.inbox.emails(:unread).each do |email|
    a = AccoutItem.new email
    worksheet[row_index, 1] = a.year
    worksheet[row_index, 2] = a.month
    worksheet[row_index, 3] = a.day
    worksheet[row_index, 4] = a.amount
    worksheet[row_index, 5] = a.place
    worksheet[row_index, 6] = a.detail
    row_index = row_index + 1

    email.mark(:read)

    puts a.to_s
  end
end

worksheet.save
worksheet.synchronize

