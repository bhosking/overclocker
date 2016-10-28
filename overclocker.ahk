#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
CoordMode, Pixel, Client ; Better portability
CoordMode, Mouse, Client ; Better portability

global msi_name := "MSI Afterburner"
global msi_graph_name := "Voltage/Frequency curve editor"
global msi_path := "C:\Program Files (x86)\MSI Afterburner\MSIAfterburner.exe"

global msi_yellow := 0x008080
global msi_node_grey := 0x7F7F7F
global msi_node_selected := 0xFFFFFF
global msi_fan_red := 0x06069E
global msi_fan_selected := 0x1F1FB7

global msi_graph_top_margin := 32
global msi_graph_bottom_margin := 41
global msi_graph_left_margin := 32
global msi_graph_right_margin := 33

global msi_fan_x := 283
global msi_fan_y := 305

global msi_curve_button_x := 291
global msi_curve_button_y := 219

global msi_profile_4_x := 638
global msi_profile_4_y := 395

global msi_profile_5_x := 666
global msi_profile_5_y := 395

global msi_save_x := 693
global msi_save_y := 395

global msi_apply_x := 469
global msi_apply_y := 339


global 3dmark_name := "3DMark Basic Edition"
global 3dmark_path := "C:\Program Files\Futuremark\3DMark\3DMark.exe"
global 3dmark_workload_name = "3DMark Workload"

global 3dmark_orange := 0x167DFD
global 3dmark_grey := 0xFBFBFB


global 3dmark_pass_x1 := 630
global 3dmark_pass_x2 := 670
global 3dmark_pass_y := 210

global 3dmark_fail_x1 := 670
global 3dmark_fail_x2 := 700
global 3dmark_fail_y := 210

global 3dmark_user_cancel_x := 1333
global 3dmark_user_cancel_y := 91

global 3dmark_benchmarks_x := 910
global 3dmark_benchmarks_y := 23

global 3dmark_firestrike_x := 1119
global 3dmark_firestrike_y := 715

global 3dmark_benchrun_x := 1197
global 3dmark_benchrun_y := 311

global 3dmark_benchmark_seconds := 428

global powershell_name := "Windows PowerShell"

global stopscript = 0

start_powershell()
sleep_or_exit(20)
main()

main() {
  WinClose, %msi_graph_name%
  Loop {
    if (get_setting("State", "Complete", 0)) {
      logi("Overclocking process is marked complete, closing script.")
      ExitApp
    }
    if (get_setting("State", "Unstable", 0))
      adjust_clock("Down")
    else if (get_setting("State", "Stable", 0))
      adjust_clock("Up")
    run_iterations()
  }
}

run_iterations() {
  start_app(msi_name, msi_path)
  Loop {
    WinActivate, %msi_name%
    PixelGetColor, fan, %msi_fan_x%, %msi_fan_x%
    if (fan <> msi_fan_red) and (fan <> msi_fan_selected) {
      click %msi_fan_x%, %msi_fan_y%
      sleep_or_exit(1)
    } else
      Break
  }
  WinActivate, %msi_name%
  click %msi_profile_5_x%, %msi_profile_5_y%
  sleep_or_exit(1)
  WinActivate, %msi_name%
  click %msi_apply_x%, %msi_apply_y%
  iteration := get_setting("State", "Iteration", 0)
  Loop {
    iteration := iteration + 1
    logi("Starting iteration " . iteration)
    pass := get_stress_test_stability()
    if (pass) {
      logi("Successfully completed iteration " . iteration)
      if (get_setting("State", "Rising", 1)) {
        if (iteration >= get_setting("UserSettings", "IterationsForStableQuick", 1)) {
          IniWrite, 1, overclocker.ini, State, Stable
          IniWrite, 0, overclocker.ini, State, Iteration
          Return
        }
      } else if (iteration >= get_setting("UserSettings", "IterationsForStableFinal", 1)) {
        logi("Successfully completed " . iteration . " iterations. Clock is considered stable.")
        WinClose, %msi_graph_name%
        WinActivate, %msi_name%
        click %msi_curve_button_x%, %msi_curve_button_y%
        sleep_or_exit(2)
        WinActivate, %msi_graph_name%
        SendInput L  ;remove lock
        WinClose, %msi_graph_name%
        sleep_or_exit(1)
        WinActivate, %msi_name%
        click %msi_save_x%, %msi_save_y%
        sleep_or_exit(1)
        WinActivate, %msi_name%
        click %msi_profile_4_x%, %msi_profile_4_y%
        logi("Saved current stable clock to profile 4.")
        start_node := get_setting("UserSettings", "HighNode", 30)
        node := get_setting("State", "Node", start_node)
        if (node <= get_setting("UserSettings", "LowNode", 1)) {
          logi("Overclock tuning is complete!")
          IniWrite, 1, overclocker.ini, State, Complete
          Shutdown, 1
        }
        node := node - 1
        IniWrite, 1, overclocker.ini, State, Rising
        IniWrite, %node%, overclocker.ini, State, Node
        IniWrite, 1, overclocker.ini, State, Stable
        IniWrite, 0, overclocker.ini, State, Iteration
        Return
      } else {
        IniWrite, %iteration%, overclocker.ini, State, Iteration
      }
    } else {
      ;sleep in case the logs haven't been read yet
      sleep_or_exit(10)
      TrayTip, INTERRUPTED, Will raise Unknown Error if script is not paused within 10 seconds, 10, 2
      sleep_or_exit(10)
      loge("Unknown Error on iteration " . iteration)
    }
  }
}


get_stress_test_stability() {  ;return 0 for unknown, 1 for stable, and should restart during sleeps if unstable
  start_app(3dmark_name, 3dmark_path)
  click %3dmark_benchmarks_x%, %3dmark_benchmarks_y%
  sleep_or_exit(1)
  WinActivate, %3dmark_name%
  click %3dmark_firestrike_x%, %3dmark_firestrike_y%
  sleep_or_exit(1)
  WinActivate, %3dmark_name%
  click %3dmark_benchrun_x%, %3dmark_benchrun_y%
  pass := 0
  Loop {
    if (!WinExist(3dmark_workload_name)) {
      if (A_Index > 60) {
        loge("Benchmark did not start")
      }
      sleep_or_exit(1)
    } else {
      break
    }
  }
  Loop {
    if (!WinExist(3dmark_workload_name)) {
      sleep_or_exit(10)
      if (!WinExist(3dmark_workload_name)) {
        if !WinActive(3dmark_name) {
          return 0
        } else {
          break
        }
      }
    }
    sleep_or_exit(1)
  }
  sleep_or_exit(20)
  WinActivate, %3dmark_name%
  PixelSearch, dummy_x, dummy_y, %3dmark_pass_x1%, %3dmark_pass_y%, %3dmark_pass_x2%, %3dmark_pass_y%, %3dmark_orange%, 0, Fast
  if (ErrorLevel = 0) {
    PixelSearch, dummy_x, dummy_y, %3dmark_pass_x1%, %3dmark_pass_y%, %3dmark_pass_x2%, %3dmark_pass_y%, %3dmark_grey%, 0, Fast
    if (ErrorLevel = 0) 
      pass := 1
  }
  return pass
}

adjust_clock(direction) {
  if direction = Down
    distance := get_setting("UserSettings", "DroppingIncrements", 12)
  else if direction = Up
    distance := get_setting("UserSettings", "ClimbingIncrements", 24)
  else
    loge("Invalid clock adjustment direction: " . direction)
  start_node := get_setting("UserSettings", "HighNode", 30)
  node := get_setting("State", "Node", start_node)
  start_app(msi_name, msi_path)
  click %msi_profile_5_x%, %msi_profile_5_y%
  sleep_or_exit(1)
  WinClose, %msi_graph_name%
  WinActivate, %msi_name%
  click %msi_curve_button_x%, %msi_curve_button_y%
  sleep_or_exit(2)
  logi("Searching for node " . node . "...")
  WinActivate, %msi_graph_name%
  WinGetPos, X, Y, width, height, %msi_graph_name%
  column := msi_graph_left_margin
  msi_graph_bottom := height - msi_graph_bottom_margin
  msi_graph_right := width - msi_graph_right_margin
  SendInput L  ;remove lock
  Loop, %node% {
    WinActivate, %msi_graph_name%
    columns := msi_graph_right - column + 1
    Loop %columns% {
      PixelSearch, node_x, node_y, %column%, %msi_graph_top_margin%, %column%, %msi_graph_bottom%, %msi_node_grey%, 0, Fast
      if ErrorLevel {
        column := column + 1
      } else {
        break
      }
    }
    if ErrorLevel
      loge("Could not find node " . A_Index)
    column := column + 6
    sleep_or_exit(0)
  }
  node_x := node_x + 2
  node_y := node_y + 2
  WinActivate, %msi_graph_name%
  click %node_x%, %node_y%
  sleep_or_exit(1)
  WinActivate, %msi_graph_name%
  PixelGetColor, color_target, %node_x%, %node_y%
  if not (color_target = msi_node_selected)
    loge("Could not click node " . node)
  logi("Adjusting node " . node . " " . direction . " " . distance . "MHz...")  
  WinActivate, %msi_graph_name%
  SendInput {%direction% %distance%}
  sleep_or_exit(1)
  WinActivate, %msi_graph_name%
  SendInput L ;reapply lock
  sleep_or_exit(1)
  WinActivate, %msi_graph_name%
  PixelSearch, dummy_x, dummy_y, node_x - 8, %node_y%, node_x + 8, node_y + 20, %msi_yellow%, 0, Fast
  if ErrorLevel
    loge("Could not lock selection")
  WinClose, %msi_graph_name%
  sleep_or_exit(1)
  WinActivate, %msi_name%
  click %msi_save_x%, %msi_save_y%
  sleep_or_exit(1)
  WinActivate, %msi_name%
  click %msi_profile_5_x%, %msi_profile_5_y%
  IniWrite, 0, overclocker.ini, State, Unstable
  IniWrite, 0, overclocker.ini, State, Stable
  IniWrite, 0, overclocker.ini, State, Iteration
  logi("Done.")
}

logi(msg) {
  FormatTime, timestamp,, yyyy-MM-dd HH:mm:ss
  entry = %timestamp% - [INFO] - %msg%`n
  FileAppend, %entry%, overclocker.log
  TrayTip, INFO, %msg%, 4, 1
}

loge(msg) {
  if get_setting("UserSettings", "ScreenshotOnError", 1) {
    SendInput {RWin Down}{PrintScreen}{RWin Up}
  }
  FormatTime, timestamp,, yyyy-MM-dd HH:mm:ss
  entry = %timestamp% - [ERROR] - %msg%, aborting.`n
  FileAppend, %entry%, overclocker.log
  TrayTip, ERROR, %msg%`, aborting., 4, 3
  close_apps()
  if get_setting("UserSettings", "RestartOnError", 0) {
    Shutdown, 2
  }
  Exit
}

run_then_close() {
  run_iterations()
  close_apps()
}

sleep_or_exit(seconds) {
  check_exit()
  loop % seconds * 2 {
    sleep 500
    FileRead, crashed, crashed.dat
    if crashed {
      FileDelete, crashed.dat
      FileAppend, 0, crashed.dat
      logi("Unstable Clock")
      IniWrite, 1, overclocker.ini, State, Unstable
      IniWrite, 0, overclocker.ini, State, Rising
      IniWrite, 0, overclocker.ini, State, Iteration
      logi("Restarting...")
      Shutdown, 2
    }
    check_exit()
  }
}

check_exit() {
  if (stopscript = 1) {
    logi("Received exit signal, stopping...")
    SendInput {Esc} ;Close 3dmark workload if running
    close_apps()
    Exit
  }
}

close_apps() {
  WinClose, %3dmark_name%
  WinClose, %msi_name%
  WinClose, %powershell_name%
}

start_app(app_title, path) {
  If !WinExist(app_title) {
    logi("Starting " . app_title . "...")
    run %path%
    Loop {
      If !WinExist(app_title) {
        if (A_Index > 60)
          loge(app_title . " took too long")
        sleep_or_exit(1)
      } else
        Break
    }
  }
  sleep_or_exit(1)
  WinActivate, %app_title%
}

get_setting(section, key, default) {
  IniRead, x, overclocker.ini, %section%, %key%, %default%
  IniWrite, %x%, overclocker.ini, %section%, %key%
  Return %x%
}

start_powershell() {
  Run C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -file "restart_on_crash.ps1"
  WinMinimize %powershell_name%
}

^p::
if (stopscript) {
  stopscript = 0
  start_powershell()
  main()
} else {
  stopscript = 1
}
Return