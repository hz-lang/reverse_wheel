use std::{collections::HashMap, io};

use anyhow::Ok;
use registry::{Data, Hive, Security};
use wmi::{COMLibrary, Variant, WMIConnection};

fn main() {
    if let Err(e) = execute() {
        println!("{e}");
    }
}

fn execute() -> anyhow::Result<()> {
    let ids = get_mouse_ids()?;
    let keys = get_keys(&ids)?;
    for (i, k) in keys.iter().enumerate() {
        let value = is_reversed(k)?;
        if value {
            println!("\n[{i}] 方向：自然");
        } else {
            println!("\n[{i}] 方向：默认");
        }

        if enter_yes()? {
            reverse(k, value)?;

            let d = if value { "默认" } else { "自然" };
            println!("反转后，当前滚动方向为：{}", d);
        }
    }
    Ok(())
}

fn enter_yes() -> anyhow::Result<bool> {
    let mut r = false;
    let mut buf = String::new();
    loop {
        println!("确定要反转滚轮方向吗？(Y/n)");

        io::stdin().read_line(&mut buf)?;
        let input = buf.as_bytes();
        match input[0] {
            b'n' => break,
            b'Y' => {
                r = true;
                break;
            }
            _ => buf.clear(),
        }
    }

    Ok(r)
}

/// 反转滚轮方向。
fn reverse(key: &str, value: bool) -> anyhow::Result<()> {
    let regkey = Hive::LocalMachine.open(key, Security::Write)?;
    let v = if value { 0 } else { 1 };
    regkey.set_value("FlipFlopWheel", &Data::U32(v))?;
    Ok(())
}

/// 滚轮方向是否已经反转。
fn is_reversed(key: &str) -> anyhow::Result<bool> {
    let regkey = Hive::LocalMachine.open(key, Security::Read)?;
    let data = regkey.value("FlipFlopWheel")?;
    match data {
        Data::U32(v) => Ok(v == 1),
        _ => Err(anyhow::anyhow!("FlipFlopWheel = {:?}", data)),
    }
}

/// 提取鼠标的设备编号。
fn get_keys(ids: &[HashMap<String, Variant>]) -> anyhow::Result<Vec<String>> {
    if ids.len() > 0 {
        println!("共查到 {} 个鼠标", &ids.len());
    }
    let mut buf = Vec::with_capacity(ids.len());
    for (i, h) in ids.iter().enumerate() {
        if let Some(val) = h.get("Description") {
            if let Variant::String(s) = val {
                println!("[{i}] {}", s);
            }
        }
        if let Some(val) = h.get("DeviceID") {
            if let Variant::String(s) = val {
                let mut k = r"SYSTEM\CurrentControlSet\Enum\".to_owned();
                k.push_str(s);
                k.push_str(r"\Device Parameters");
                buf.push(k);
            }
        }
    }
    Ok(buf)
}

/// 获取鼠标的描述信息和设备编号。
fn get_mouse_ids() -> anyhow::Result<Vec<HashMap<String, Variant>>> {
    let com_con = COMLibrary::new()?;
    let wmi_con = WMIConnection::new(com_con.into())?;

    let cmd = "SELECT Description, DeviceID \
                     FROM Win32_PnPEntity \
                     WHERE ClassGuid = \"{4d36e96f-e325-11ce-bfc1-08002be10318}\"";
    let r: Vec<HashMap<String, Variant>> = wmi_con.raw_query(cmd)?;
    Ok(r)
}
