import logging
from datetime import datetime, timedelta
from pandas import DataFrame
from slack_sdk import WebClient

class SlackTradesExporter:
    def __init__(self, token, channel_id):
        self.client = WebClient(token=token)
        self.channel_id = channel_id

    def fetch_messages(self, oldest_timestamp, limit=1000):
        result = self.client.conversations_history(
            channel=self.channel_id,
            inclusive=True,
            oldest=oldest_timestamp,
            limit=limit
        )
        return result["messages"]

    def parse_message(self, message):
        text = message["text"]
        try:
            timestamp = float(message["ts"])
            dt = datetime.utcfromtimestamp(timestamp) + timedelta(hours=3)
            date = dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception as e:
            date = None
            timestamp = None
            logging.error(f"Error parsing timestamp for message: {e}")

        fields = {
            "Action": self._extract_value("Action: ", text),
            "Entry Price": self._extract_value("Entry Price: ", text),
            "Close Price": self._extract_value("Close Price: ", text),
            "Stop Loss Price": self._extract_value("Stop Loss Price: ", text),
            "Take Profit Price": self._extract_value("Take Profit Price: ", text),
            "Position Size": self._extract_value("Position Size: ", text),
            "Total Position Size": self._extract_value("Total Position Size: ", text),
            "Closed Position Size": self._extract_value("Closed Position Size: ", text),
            "Number of Positions": self._extract_value("Number of Positions: ", text),
            "Date": date
        }
        return fields

    def _extract_value(self, prefix, text):
        try:
            return text.split(prefix)[1].split("\n")[0]
        except IndexError:
            return None

    def create_dataframe(self, messages):
        data = [self.parse_message(message) for message in messages]
        df = DataFrame(data)
        return df.sort_values(by="Date")

    def export_to_excel(self, df, filename):
        df.to_excel(filename, index=False)


if __name__ == "__main__":
    token = "SLACK TOKEN"
    channel_id = "CHANNEL ID"
    oldest_timestamp = "1704056400"

    exporter = SlackTradesExporter(token, channel_id)
    messages = exporter.fetch_messages(oldest_timestamp)
    df = exporter.create_dataframe(messages)
    exporter.export_to_excel(df, "trades.xlsx")