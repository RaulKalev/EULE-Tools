import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System.Net')

from System.Drawing import Size, Point
from System.Windows.Forms import Application, Form, TextBox, Button, ScrollBars, OpenFileDialog, FormBorderStyle, DialogResult, FormStartPosition, Label
from System.Net import WebClient
import json
import os

class InputForm(Form):
    def __init__(self, title, label_text):
        self.Text = title
        self.Size = Size(300, 120)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterParent

        self.label = Label()
        self.label.Text = label_text
        self.label.Location = Point(10, 10)
        self.label.Size = Size(280, 20)
        self.Controls.Add(self.label)

        self.textBox = TextBox()
        self.textBox.Location = Point(10, 40)
        self.textBox.Size = Size(280, 20)
        self.Controls.Add(self.textBox)

        self.okButton = Button()
        self.okButton.Text = "OK"
        self.okButton.Location = Point(210, 70)
        self.okButton.Click += self.on_ok_click
        self.Controls.Add(self.okButton)

        self.AcceptButton = self.okButton

    def on_ok_click(self, sender, event):
        self.DialogResult = DialogResult.OK

class ChatForm(Form):
    api_key = None
    api_key_file_path = "api_key_path.txt"  # File to store the API key file path

    def __init__(self):
        self.Text = "Chat with GPT"
        self.Size = Size(500, 300)

        if not self.load_api_key_path():
            self.select_api_key_file()

        # Output Box
        self.outputBox = TextBox()
        self.outputBox.Multiline = True
        self.outputBox.ScrollBars = ScrollBars.Vertical
        self.outputBox.ReadOnly = True
        self.outputBox.Location = Point(10, 10)
        self.outputBox.Size = Size(480, 200)
        self.Controls.Add(self.outputBox)

        # Input Box
        self.inputBox = TextBox()
        self.inputBox.Location = Point(10, 220)
        self.inputBox.Size = Size(400, 20)
        self.Controls.Add(self.inputBox)

        # Send Button
        self.sendButton = Button()
        self.sendButton.Text = 'Send'
        self.sendButton.Location = Point(420, 220)
        self.sendButton.Click += self.on_send_click
        self.Controls.Add(self.sendButton)

        self.AcceptButton = self.sendButton

        if ChatForm.api_key is None:
            self.select_api_key_file()

    def select_api_key_file(self):
        openFileDialog = OpenFileDialog()
        openFileDialog.Filter = "Text files (*.txt)|*.txt"
        openFileDialog.Title = "Select API Key File"

        if openFileDialog.ShowDialog() == DialogResult.OK:
            with open(openFileDialog.FileName, 'r') as file:
                ChatForm.api_key = file.read().strip()
            self.save_api_key_path(openFileDialog.FileName)

    def save_api_key_path(self, path):
        with open(ChatForm.api_key_file_path, 'w') as file:
            file.write(path)

    def load_api_key_path(self):
        try:
            with open(ChatForm.api_key_file_path, 'r') as file:
                key_path = file.read().strip()
                if os.path.isfile(key_path):
                    with open(key_path, 'r') as key_file:
                        ChatForm.api_key = key_file.read().strip()
                    return True
        except IOError:
            return False
        return False

    def on_send_click(self, sender, event):
        user_input = self.inputBox.Text
        self.inputBox.Clear()
        new_message = "You: " + user_input + "\r\n"
        self.outputBox.Text = new_message + self.outputBox.Text

        response = self.query_chatgpt(user_input)
        self.outputBox.Text = "GPT: " + response + "\r\n" + self.outputBox.Text

    def query_chatgpt(self, prompt):
        if ChatForm.api_key is None:
            return "API key not set."
        api_url = "https://api.openai.com/v1/chat/completions"
        client = WebClient()

        client.Headers.Add("Authorization", "Bearer " + ChatForm.api_key)
        client.Headers.Add("Content-Type", "application/json")

        data = {
            "model": "gpt-3.5-turbo",
            "messages": [{"role": "system", "content": "You are a helpful assistant."},
                         {"role": "user", "content": prompt}]
        }

        try:
            response = client.UploadString(api_url, json.dumps(data))
            response_json = json.loads(response)
            return response_json['choices'][0]['message']['content']
        except Exception as e:
            return "Error: " + str(e)

Application.EnableVisualStyles()
form = ChatForm()
form.ShowDialog()
