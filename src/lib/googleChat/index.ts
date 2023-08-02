/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
const env = {
  googleChat: PropertiesService.getScriptProperties().getProperty(
    'GOOGLE_CHAT_WEBHOOK'
  ),
};

const config = (text: string) => ({
  method: 'post' as GoogleAppsScript.URL_Fetch.HttpMethod,
  headers: {
    'Content-Type': 'application/json; charset=UTF-8',
  },
  payload: JSON.stringify({ text }),
});

type Prop = {
  text: string;
};

/**
 * google chatにメッセージを送信する
 * @param {{text}} string メッセージテキスト
 */
function sendGoogleChat(prop: Prop = { text: 'Hello World!' }) {
  const { text } = prop;

  if (env.googleChat) {
    const res = UrlFetchApp.fetch(env.googleChat, config(text));

    console.log('done');
  } else {
    throw new Error('google chatのwebhookが指定されていません');
  }
}
