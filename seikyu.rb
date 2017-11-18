require 'google/apis/calendar_v3'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'time'
require 'date'
require 'pry'
require 'fileutils'
require 'Spreadsheet'
require 'active_support/all'

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
APPLICATION_NAME = 'Google Calendar API '
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials',
                             "calendar-ruby-quickstart.yaml")
SCOPE = Google::Apis::CalendarV3::AUTH_CALENDAR_READONLY

def authorize
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))

  client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(
    client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(
      base_url: OOB_URI)
    puts "Open the following URL in the browser and enter the " +
         "resulting code after authorization"
    puts url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI)
  end
  credentials
end
puts "Please input year"
year = gets.chomp.to_s
puts "Please input month"
month = gets.chomp.to_i
if month == 12
  next_month = 1
else
  next_month  = month+1
end
puts "Please input project name"
project = gets.chomp.to_s

puts "Please input wages per hour"
wage = gets.chomp.to_i


time_min = "#{year}-#{month}-01 00:00:00"
time_max = "#{year}-#{next_month}-01 00:00:00"

# Initialize the API
service = Google::Apis::CalendarV3::CalendarService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

# Fetch the next 10 events for the user
calendar_id = 'primary'
response = service.list_events(calendar_id,
                               max_results: 2500,
                               single_events: true,
                               time_min: Time.parse(time_min).iso8601,
                               time_max: Time.parse(time_max).iso8601,
                               order_by: 'startTime')
total_spend_time = 0
response.items.each do |event|
  if event.summary.nil? || event.start.date_time.nil? then
    next
  end
  start = event.start.date || event.start.date_time
  spend_time = ((event.end.date_time.-event.start.date_time)*24).to_f
  if event.summary.include?(project)
    puts "- #{event.summary} (#{start})(#{spend_time}時間)"
    total_spend_time += spend_time
  end
end
puts "Total project hours are #{total_spend_time} hours"
total_fee = total_spend_time * wage
str_total_spend_time = total_spend_time.to_s
puts "Totall fee is  #{total_fee} yen"

book = Spreadsheet.open('invoice.xls')
sheet = book.worksheet('請求書')
sheet[13,8] = total_fee
sheet[16,2] = "業務委託費_#{month}月分"
sheet[16,15] = str_total_spend_time
sheet[16,18] = wage
sheet[16,21] = total_fee

today_month = Time.now.month.to_s
today_day = Time.now.day.to_s

sheet[5,20] = today_month
sheet[5,22] = today_day

sheet[30,0] = "平成29年#{Time.now.end_of_month.month}月#{Time.now.end_of_month.day}日までに下記振込先にお振込みをお願いいたします。"


book.write "請求書_橘_#{month}月分.xls"
