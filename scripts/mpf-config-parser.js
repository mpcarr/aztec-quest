const fs = require('fs');
const yaml = require('js-yaml');

// Function to read and parse the YAML file
function readYamlFile(filePath) {
  try {
    const fileContents = fs.readFileSync(filePath, 'utf8');
    return yaml.load(fileContents);
  } catch (e) {
    console.error(`Error reading YAML file: ${e}`);
    return null;
  }
}

function mapDevicesToVbs(configObject) {
  let vbsCode = `\n'Devices\n`;
  if (configObject.ball_devices) {
    const { switches, ball_saves, playfields, ball_devices } = configObject;
    let defaultSourceDevice = null;
    if (playfields) {
      defaultSourceDevice = playfields.playfield.default_source_device
    }
    console.log(defaultSourceDevice)

    if (ball_devices) {
      for (const ballDevice of Object.keys(ball_devices)) {
        if (!ball_devices[ballDevice].tags || ball_devices[ballDevice].tags.indexOf("trough") == -1) {
          vbsCode += `Dim ${ballDevice}: Set ${ballDevice} = (new BallDevice)("${ballDevice}", "${ball_devices[ballDevice].ball_switches}", ${ball_devices[ballDevice].player_controlled_eject_events || 'Null'}, ${ball_devices[ballDevice].eject_timeouts.replace("s", "")}, ${ballDevice == defaultSourceDevice ? "True" : "False"}, ${ball_devices[ballDevice].debug || "False"})`
          vbsCode += '\n';
        }
      }
    }

  }
  return vbsCode;
}

function mapDivertersToVbs(configObject) {
  let vbsCode = `\n'Diverters\n`;
  if (configObject.switches) {
    const { diverters } = configObject;

    for (const d of Object.keys(diverters)) {
      console.log(diverters[d]);
      vbsCode += `Dim ${d} : Set ${d} = (new Diverter)("${d}", Array(${diverters[d].enable_events.split(",").map(item => `"${item}"`).join(",")}), Array(${diverters[d].disable_events.split(",").map(item => `"${item}"`).join(",")}), Array(${diverters[d].activate_events.split(",").map(item => `"${item}"`).join(",")}), Array(${diverters[d].deactivate_events.split(",").map(item => `"${item}"`).join(",")}), ${diverters[d].activation_time}, ${diverters[d].debug || "False"})`
      vbsCode += '\n';
    }
  }

  return vbsCode;
}

function mapDropTargetsToVbs(configObject) {
  let vbsCode = `\n'Drop Targets\n`;
  if (configObject.switches) {
    const { drop_targets } = configObject;

    for (const d of Object.keys(drop_targets)) {
      console.log(drop_targets[d]);
      vbsCode += `Dim ${d} : Set ${d} = (new DropTarget)(${drop_targets[d].switch}, ${drop_targets[d].switch}a, BM_${drop_targets[d].switch}, ${extractNumberWithoutPadding(drop_targets[d].switch)}, 0, False, Array(${drop_targets[d].reset_events.split(",").map(item => `"${item}"`).join(",")}))`
      vbsCode += '\n';
    }
  }

  return vbsCode;
}


function extractNumberWithoutPadding(str) {
  return str.replace(/[^\d]/g, '').replace(/^0+/, '');
}


function mapSwitchesToVbs(configObject) {
  let vbsCode = `\n'Switches\n`;
  if (configObject.switches) {
    const { switches } = configObject;

    for (const sw of Object.keys(switches)) {
      if(switches[sw].tags && switches[sw].tags.indexOf("trough")>-1)
        continue;

      if(switches[sw].tags && switches[sw].tags.indexOf("flipper")>-1)
        continue;

      vbsCode += `
Sub ${sw}_Hit()   : DispatchPinEvent "${sw}_active",   ActiveBall : End Sub
Sub ${sw}_Unhit() : DispatchPinEvent "${sw}_inactive", ActiveBall : End Sub
`
    }
  }
  fs.writeFile('./src/game/switches/_all.switches-vpx.vbs', vbsCode, (err) => {
    if (err) {
      console.error(`Error writing VBS file: ${err}`);
    } else {
      console.log('VBS code was successfully saved to switches.vbs');
    }
  });

  return vbsCode;
}

// Main function to process the YAML file and generate VBS code
function processYamlToVbs(filePath) {
  const yamlObject = readYamlFile(filePath);
  if (yamlObject) {
    let devicesVbs = ''
    devicesVbs+= mapDevicesToVbs(yamlObject);
    devicesVbs+= mapDivertersToVbs(yamlObject);
    devicesVbs+= mapDropTargetsToVbs(yamlObject);
    fs.writeFile('./src/game/_logicStartDevices-vpx.vbs', devicesVbs, (err) => {
      if (err) {
        console.error(`Error writing VBS file: ${err}`);
      } else {
        console.log('VBS code was successfully saved to _logicStartDevices.vbs');
      }
    });
    //mapSwitchesToVbs(yamlObject);

    // Here you could further write this VBS code to a .vbs file or process it as needed
  }
}

// Replace 'your_mpf_config.yaml' with the path to your actual MPF YAML configuration file
const filePath = '../mpf/config/config.yaml';
processYamlToVbs(filePath);
