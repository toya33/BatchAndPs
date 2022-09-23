# 事前にpowershellで「Get-AudioDevice -List」で使用したい
# サウンド(Type:Playback)とマイク(Type:Recording)のIDを
# 下記の配列に設定しておく
# (スピーカー, イヤホン)
$SoundDeviceIDList = @("<ID>");
$RecordingDeviceIDList = @("<ID>");

function Get-NextSoundeDeviceID {
    param (
        $CurrentSoundDeviceID,
        $CurrentRecordingDeviceID
    )

    
    
    # サウンド選択
    if (  $SoundDeviceIDList[$SoundDeviceIDList.GetUpperBound(0)] -eq $CurrentSoundDeviceID ) {
        # 現在のサウンドデバイスにサウンドデバイスリストの末尾が選択されていた場合
        # 次のサウンドデバイスにはサウンでデバイスリストの先頭を設定する
        $NextSoundDeviceID = $SoundDeviceIDList[$SoundDeviceIDList.GetLowerBound(0)];
    }
    else {
        # リストの次の値を設定する
        for ( $index = 0; $index -lt $SoundDeviceIDList.count; $index++) {
            if ( $CurrentSoundDeviceID -eq $SoundDeviceIDList[$index] ) {
                $NextSoundDeviceID = $SoundDeviceIDList[$index + 1];
            }
        }
    }

    # マイク選択
    if (  $RecordingDeviceIDList[$RecordingDeviceIDList.GetUpperBound(0)] -eq $CurrentRecordingDeviceID ) {
        $NextRecordingDeviceID = $RecordingDeviceIDList[$RecordingDeviceIDList.GetLowerBound(0)];
    }
    else {
        # リストの次の値を設定する
        for ( $index = 0; $index -lt $RecordingDeviceIDList.count; $index++) {
            if ( $CurrentRecordingDeviceID -eq $RecordingDeviceIDList[$index] ) {
                $NextRecordingDeviceID = $RecordingDeviceIDList[$index + 1];
            }
        }
    }

    return $NextSoundDeviceID, $NextRecordingDeviceID
}

function Set-SoundAndRecordingDeviceID {
    param (
        $SoundDeviceID,
        $RecordingDeviceID
    )

    #TODO:出力先が存在するかどうか判定する
    Set-AudioDevice -ID $SoundDeviceID;
    Set-AudioDevice -ID $RecordingDeviceID;

}

function Push-SoundAndRecordingDeviceInfoToToast {
    param (
        $SoundDeviceID,
        $RecordingDeviceID
    )

    $AppId = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe";
    $template =
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::GetTemplateContent(
        [Windows.UI.Notifications.ToastTemplateType, Windows.UI.Notifications, ContentType = WindowsRuntime]::ToastText01);
    $msg = "Playback device : `n" + (Get-AudioDevice -ID $SoundDeviceID).Name + "`n";
    $msg += "Recording device : `n" + (Get-AudioDevice -ID $RecordingDeviceID).Name;
    $template.GetElementsByTagName("text").Item(0).InnerText = $msg;
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId).Show($template);

}

Get-AudioDevice -Playback
Get-AudioDevice -Recording

# スリープしている間に変える必要がなければ「Ctrl + c」でキャンセルする
Start-Sleep 2

$currentSoundDeviceID = (Get-AudioDevice -Playback).ID;
$currentRecordingDeviceID = (Get-AudioDevice -Recording).ID;

$nextDevice = Get-NextSoundeDeviceID -CurrentSoundDeviceID $currentSoundDeviceID $currentRecordingDeviceID;

$nextDevice

Set-SoundAndRecordingDeviceID $nextDevice[0] $nextDevice[1]

Start-Sleep 1

Push-SoundAndRecordingDeviceInfoToToast -SoundDeviceID $nextDevice[0] $nextDevice[1]