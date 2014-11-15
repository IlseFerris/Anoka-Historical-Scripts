'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""
EMSendKey "i" + "<enter>"

EMWaitReady 1, 0
EMSetCursor 20, 71
EMSendKey "sves" + "<enter>"

EMWaitReady 1, 0
EMSetCursor 20, 70
EMSendKey "tpqy" + "<enter>"
