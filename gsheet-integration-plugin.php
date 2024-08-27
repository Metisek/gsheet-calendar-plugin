<?php
/*
Plugin Name: GSheet Integration Plugin
Description: A plugin to integrate with Google Sheets, allowing viewing and editing of spreadsheet data.
Version: 1.0
Author: Your Name
*/

defined('ABSPATH') || exit;

// Include the Google Client library
require_once plugin_dir_path(__FILE__) . 'lib/google-client.php';

define('GIP_PLUGIN_NAME', __('GSheet Integration Plugin', 'gsheet-integration-plugin'));
define('GIP_OAUTH_OPTION_NAME', 'gip_google_oauth');
define('GIP_SPREADSHEET_ID_OPTION_NAME', 'gip_spreadsheet_id');
define('GIP_SHEET_ID_OPTION_NAME', 'gip_sheet_id');

// Register plugin settings page
add_action('admin_menu', 'gip_register_settings_page');
function gip_register_settings_page() {
    add_options_page(
        __('GSheet Integration Settings', 'gsheet-integration-plugin'),
        __('GSheet Integration', 'gsheet-integration-plugin'),
        'manage_options',
        'gip',
        'gip_settings_page_html'
    );
}

// Settings page HTML
function gip_settings_page_html() {
    if (!current_user_can('manage_options')) {
        return;
    }

    // Save settings
    if (isset($_POST['gip_save_settings'])) {
        update_option(GIP_OAUTH_OPTION_NAME, sanitize_text_field($_POST['gip_oauth']));
        update_option(GIP_SPREADSHEET_ID_OPTION_NAME, sanitize_text_field($_POST['gip_spreadsheet_id']));
        update_option(GIP_SHEET_ID_OPTION_NAME, sanitize_text_field($_POST['gip_sheet_id']));
        echo '<div class="updated"><p>' . __('Settings saved.', 'gsheet-integration-plugin') . '</p></div>';
    }

    // Get current settings
    $oauth_data = get_option(GIP_OAUTH_OPTION_NAME);
    $spreadsheet_id = get_option(GIP_SPREADSHEET_ID_OPTION_NAME);
    $sheet_id = get_option(GIP_SHEET_ID_OPTION_NAME);

    ?>
    <div class="wrap">
        <h1><?php echo esc_html(get_admin_page_title()); ?></h1>
        <form method="post" action="">
            <?php wp_nonce_field('gip_save_settings', 'gip_nonce'); ?>
            <table class="form-table">
                <tr valign="top">
                    <th scope="row"><?php _e('OAuth Client Data (JSON)', 'gsheet-integration-plugin'); ?></th>
                    <td><textarea name="gip_oauth" rows="10" cols="50" class="large-text"><?php echo esc_textarea($oauth_data); ?></textarea></td>
                </tr>
                <tr valign="top">
                    <th scope="row"><?php _e('Spreadsheet ID', 'gsheet-integration-plugin'); ?></th>
                    <td><input type="text" name="gip_spreadsheet_id" value="<?php echo esc_attr($spreadsheet_id); ?>" class="regular-text"></td>
                </tr>
                <tr valign="top">
                    <th scope="row"><?php _e('Sheet ID (GID)', 'gsheet-integration-plugin'); ?></th>
                    <td><input type="text" name="gip_sheet_id" value="<?php echo esc_attr($sheet_id); ?>" class="regular-text"></td>
                </tr>
            </table>
            <p class="submit">
                <input type="submit" name="gip_save_settings" class="button-primary" value="<?php _e('Save Settings', 'gsheet-integration-plugin'); ?>">
            </p>
        </form>

        <h2><?php _e('Authorize and Interact with Google Sheets', 'gsheet-integration-plugin'); ?></h2>
        <form method="post" action="">
            <input type="submit" name="gip_authorize" class="button-primary" value="<?php _e('Authorize', 'gsheet-integration-plugin'); ?>">
            <input type="submit" name="gip_view_sheet" class="button-secondary" value="<?php _e('View Sheet', 'gsheet-integration-plugin'); ?>">
        </form>

        <h2><?php _e('Insert Data into Google Sheet', 'gsheet-integration-plugin'); ?></h2>
        <form method="post" action="">
            <table class="form-table">
                <tr valign="top">
                    <th scope="row"><?php _e('Row Number', 'gsheet-integration-plugin'); ?></th>
                    <td><input type="number" name="gip_row_number" class="small-text"></td>
                </tr>
                <tr valign="top">
                    <th scope="row"><?php _e('Column Number', 'gsheet-integration-plugin'); ?></th>
                    <td><input type="number" name="gip_column_number" class="small-text"></td>
                </tr>
                <tr valign="top">
                    <th scope="row"><?php _e('Data to Insert', 'gsheet-integration-plugin'); ?></th>
                    <td><input type="text" name="gip_data_to_insert" class="regular-text"></td>
                </tr>
            </table>
            <p class="submit">
                <input type="submit" name="gip_insert_data" class="button-primary" value="<?php _e('Insert Data', 'gsheet-integration-plugin'); ?>">
            </p>
        </form>
    </div>
    <?php
}

// Handle OAuth Authorization and Viewing Sheet
if (isset($_POST['gip_authorize'])) {
    // Handle the OAuth authorization process here
    gip_authorize_google_client();
}

if (isset($_POST['gip_view_sheet'])) {
    $spreadsheet_id = get_option(GIP_SPREADSHEET_ID_OPTION_NAME);
    $sheet_id = get_option(GIP_SHEET_ID_OPTION_NAME);
    if ($spreadsheet_id && $sheet_id) {
        $url = "https://docs.google.com/spreadsheets/d/$spreadsheet_id/edit#gid=" . $sheet_id;
        wp_redirect($url);
        exit;
    }
}

// Handle inserting data into the Google Sheet
if (isset($_POST['gip_insert_data'])) {
    $row = intval($_POST['gip_row_number']);
    $col = intval($_POST['gip_column_number']);
    $data = sanitize_text_field($_POST['gip_data_to_insert']);
    gip_insert_data_into_sheet($row, $col, $data);
}

// Authorize the Google Client
function gip_authorize_google_client() {
    // Implement Google OAuth authorization logic here using google-client.php
    $oauth_data = get_option(GIP_OAUTH_OPTION_NAME);
    if ($oauth_data) {
        $client = create_google_client($oauth_data);
        // Redirect user to the OAuth URL or perform authorization
    } else {
        echo '<div class="error"><p>' . __('OAuth data missing. Please enter it in the settings.', 'gsheet-integration-plugin') . '</p></div>';
    }
}

// Insert data into the Google Sheet
function gip_insert_data_into_sheet($row, $col, $data) {
    // Implement logic to insert data into the Google Sheet
    $spreadsheet_id = get_option(GIP_SPREADSHEET_ID_OPTION_NAME);
    $sheet_id = get_option(GIP_SHEET_ID_OPTION_NAME);
    if ($spreadsheet_id && $sheet_id) {
        $client = create_google_client(get_option(GIP_OAUTH_OPTION_NAME));
        $service = new Google_Service_Sheets($client);
        $range = "'$sheet_id'!R$row:C$col";
        $values = [
            [$data]
        ];
        $body = new Google_Service_Sheets_ValueRange([
            'values' => $values
        ]);
        $params = [
            'valueInputOption' => 'RAW'
        ];
        $service->spreadsheets_values->update($spreadsheet_id, $range, $body, $params);
        echo '<div class="updated"><p>' . __('Data inserted successfully.', 'gsheet-integration-plugin') . '</p></div>';
    } else {
        echo '<div class="error"><p>' . __('Spreadsheet ID or Sheet ID missing. Please enter it in the settings.', 'gsheet-integration-plugin') . '</p></div>';
    }
}
?>
