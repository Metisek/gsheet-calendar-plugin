<?php

class GoogleSheetsClient
{
    const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

    private $clientSecretData;
    private $accessTokenInfo;
    private $spreadsheetId;
    private $sheetName;

    public function __construct($clientSecretData, $spreadsheetId, $sheetName)
    {
        $this->clientSecretData = $clientSecretData;
        $this->spreadsheetId = $spreadsheetId;
        $this->sheetName = $sheetName;
        $this->loadAccessToken();
    }

    private function loadAccessToken()
    {
        $this->accessTokenInfo = get_option('gsheet_access_token_info', []);

        if (empty($this->accessTokenInfo) || $this->isAccessTokenExpired()) {
            if (!empty($this->accessTokenInfo['refresh_token'])) {
                $this->refreshAccessToken();
            } else {
                $this->authorize();
            }
        }
    }

    private function saveAccessToken($tokenInfo)
    {
        $this->accessTokenInfo = $tokenInfo;
        update_option('gsheet_access_token_info', $tokenInfo);
    }

    private function authorize()
    {
        $authUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' . http_build_query([
            'client_id' => $this->getClientId(),
            'redirect_uri' => $this->getRedirectUri(),
            'response_type' => 'code',
            'scope' => self::SCOPES,
            'access_type' => 'offline',
            'include_granted_scopes' => 'true',
            'prompt' => 'consent'
        ]);

        echo 'Open this link in your browser to authorize: <a href="' . $authUrl . '">' . $authUrl . '</a>';
        exit;
    }

    public function handleAuthorizationCallback()
    {
        if (isset($_GET['code'])) {
            $response = wp_remote_post('https://oauth2.googleapis.com/token', [
                'body' => [
                    'code' => $_GET['code'],
                    'client_id' => $this->getClientId(),
                    'client_secret' => $this->getClientSecret(),
                    'redirect_uri' => $this->getRedirectUri(),
                    'grant_type' => 'authorization_code',
                ],
            ]);

            if (is_wp_error($response)) {
                throw new Exception('Error during authorization: ' . $response->get_error_message());
            }

            $tokenInfo = json_decode(wp_remote_retrieve_body($response), true);
            $this->saveAccessToken($tokenInfo);
        } else {
            throw new Exception('Authorization code missing.');
        }
    }

    private function getClientId()
    {
        return $this->clientSecretData['web']['client_id'];
    }

    private function getClientSecret()
    {
        return $this->clientSecretData['web']['client_secret'];
    }

    private function getRedirectUri()
    {
        return admin_url('options-general.php?page=gip');
    }

    private function isAccessTokenExpired()
    {
        return $this->accessTokenInfo['expires_in'] + $this->accessTokenInfo['created'] < time();
    }

    private function refreshAccessToken()
    {
        $response = wp_remote_post('https://oauth2.googleapis.com/token', [
            'body' => [
                'client_id' => $this->getClientId(),
                'client_secret' => $this->getClientSecret(),
                'refresh_token' => $this->accessTokenInfo['refresh_token'],
                'grant_type' => 'refresh_token',
            ],
        ]);

        if (is_wp_error($response)) {
            throw new Exception('Error refreshing token: ' . $response->get_error_message());
        }

        $tokenInfo = json_decode(wp_remote_retrieve_body($response), true);
        $tokenInfo['refresh_token'] = $this->accessTokenInfo['refresh_token'];
        $tokenInfo['created'] = time(); // Set current time as the creation time
        $this->saveAccessToken($tokenInfo);
    }

    public function readAllData()
    {
        $url = 'https://sheets.googleapis.com/v4/spreadsheets/' . $this->spreadsheetId . '/values/' . $this->sheetName;
        $response = $this->doRequest($url);

        return isset($response['values']) ? $response['values'] : null;
    }

    public function writeDataToCell($cell, $data)
    {
        $url = 'https://sheets.googleapis.com/v4/spreadsheets/' . $this->spreadsheetId . '/values/' . $this->sheetName . '!' . $cell . '?valueInputOption=RAW';
        $body = json_encode(['values' => [[is_array($data) ? implode(',', $data) : $data]]]);

        $this->doRequest($url, $body, 'PUT');
    }

    private function doRequest($url, $body = null, $method = 'GET')
    {
        $headers = [
            'Authorization' => 'Bearer ' . $this->accessTokenInfo['access_token'],
            'Content-Type' => 'application/json',
        ];

        $args = [
            'headers' => $headers,
            'method' => $method,
        ];

        if ($body) {
            $args['body'] = $body;
        }

        $response = wp_remote_request($url, $args);

        if (is_wp_error($response)) {
            throw new Exception('Request failed: ' . $response->get_error_message());
        }

        return json_decode(wp_remote_retrieve_body($response), true);
    }
}
?>
