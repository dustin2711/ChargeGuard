# C# script to change volume. Thanks to https://stackoverflow.com/questions/255419/how-can-i-mute-unmute-my-sound-from-powershell/19348221#19348221
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;

[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
  // f(), g(), ... are unused COM method slots. Define these if you care
  int f(); int g(); int h(); int i();
  int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
  int j();
  int GetMasterVolumeLevelScalar(out float pfLevel);
  int k(); int l(); int m(); int n();
  int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
  int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
  int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
  int f(); // Unused
  int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

public class Audio {
  static IAudioEndpointVolume Vol() {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
    IAudioEndpointVolume epv = null;
    var epvid = typeof(IAudioEndpointVolume).GUID;
    Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
    return epv;
  }
  public static float Volume {
    get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
    set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
  }
  public static bool Mute {
    get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
    set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
  }
}
'@


# Returns true if the battery status is charging or fully charged
function DeviceIsCharging
{
  param (
      [int]$Status
  )
  return ($Status -eq 2 -or $Status -eq 3)
}

# Gets a string for the current battery status
function GetBatteryStatusString {
  param (
      [int]$Status
  )

  switch ($Status) {
      0 { return "Discharging" }
      1 { return "Idle" }
      2 { return "Charging" }
      3 { return "Fully charged" }
      4 { return "Unknown status" }
      Default { return "Uninitialized" }
  }
}

Write-Host "This script will make an alarm if the charging cable is unplugged."

# Setup audio player
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$mp3File = Join-Path -Path $scriptPath -ChildPath "Alarm.mp3"
Add-Type -AssemblyName presentationCore
$mediaPlayer = New-Object system.windows.media.mediaplayer
$mediaPlayer.open($mp3File)

$previousState = -1
$deviceWasCharging = $false

while ($true)
{
  $powerStatus = Get-WmiObject -Class Win32_Battery
  $currentState = $powerStatus.BatteryStatus
  
  if ($previousState -eq $currentState)
  {
    continue
  }

  # Write-Host "Power state changed from $previousState to $currentState"
  $previousStateString = GetBatteryStatusString -Status $previousState
  $currentStateString = GetBatteryStatusString -Status $currentState
  Write-Host "Power state changed from $previousStateString to $currentStateString"
  # Write-Host "Power state changed from $($GetBatteryStatusString.Invoke($previousState)) to $($GetBatteryStatusString.Invoke($currentState))"

  $deviceIsCharging = DeviceIsCharging -Status $currentState
    
  if (-not $deviceWasCharging -and $deviceIsCharging)
  {
    $mediaPlayer.Stop()
    Write-Host "Charging cable plugged in. Alarm is paused."
  }
  elseif ($deviceWasCharging -and -not $deviceIsCharging)
  {
      # Battery is discharging (cable removed)
      Write-Host "Charging cable removed. Playing alarm."

      # Play alarm sound
      [Audio]::Mute = $false
      [Audio]::Volume = 1.0
      $mediaPlayer.Play()
  }
  else
  {
      Write-Host "Battery is not charging."
  }
  
  $deviceWasCharging = $deviceIsCharging
  $previousState = $currentState
  Start-Sleep -Milliseconds 100

  Write-Host "Waiting for changes..."
}
