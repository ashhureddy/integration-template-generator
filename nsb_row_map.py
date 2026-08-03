"""
Row map for the real NSB_Macro_Template_v4 template (single sheet: NSB).
Every row number below is taken directly from the actual confirmed dump of that file.
"""

NSB_ROW_MAP = {
    # ---- Header / Configuration / IDL Connections ----
    "subject": 3,
    "iwm_details": 6,
    "pre_configuration": 10,       # always "N/A"
    "current_configuration": 11,
    "post_configuration": 12,
    "wll_node": 13,
    "controller_6610": 14,
    "software_version": 15,
    "gs_version": 16,
    "idle": [19, 20],
    "idly": 21,
    "switch": 23,
    "slot_port": 24,
    "manual_feed_buffer": list(range(25, 40)),

    # ---- Completed (41 = section header) ----
    "integration": {"completed": [42, 43, 44], "pending": [94, 95]},
    "controller_integration": {"completed": [45], "pending": [96]},
    "dss_activation": {"completed": [46], "pending": [97]},
    "ngs_activation": {"completed": [47], "pending": [98]},
    "gps_installation": {"completed": [48], "pending": [99]},
    "lkf_installation": {"completed": [49], "pending": [100]},
    "psap_speedtest": {"completed": [50], "pending": [101]},
    "speedtest_lte": {"completed": [51], "pending": [102]},
    "speed_test_5g": {"completed": [52], "pending": [103]},
    "calltest_fnet": {"completed": [53], "pending": [104]},
    "transport_sfp": {"completed": [54, 55, 56], "pending": [107, 108, 109]},
    "ret_configuration": {"completed": [57], "pending": [110]},
    "external_alarm_scripting": {"completed": [58], "pending": [111]},
    "sau_connections": {"completed": [59], "pending": [112]},
    "sup_connections": {"completed": [60], "pending": [113]},
    "xmu_installation": {"completed": [61], "pending": [114]},
    "idl_connections": {"completed": [62], "pending": [116]},
    "area_test": {"completed": [63], "pending": [119]},
    "external_alarm_testing": {"completed": [64], "pending": [120]},
    "script_load_6673": {"completed": [65], "pending": [117]},
    "installation_manual": {"completed": [66], "pending": None},

    # ---- Pending-only ----
    "post_configuration_pending": {"completed": None, "pending": [93]},
    "sfp_installation_bbu": {"completed": None, "pending": [105]},
    "sfp_installation_radio": {"completed": None, "pending": [106]},
    "rilinks_scripting": {"completed": None, "pending": [115]},
    "siad_provisioning": {"completed": None, "pending": [118]},
    "config_6673": {"completed": None, "pending": [121]},
    "port_config_6673_enm": {"completed": None, "pending": [122]},
    "link_failure": {"completed": None, "pending": [123]},
    "sfp_not_present": {"completed": None, "pending": [124]},
    "mo_inconsistent_config_alarm": {"completed": None, "pending": [125]},
    "fiberloss": {"completed": None, "pending": [126, 127]},   # Data Link_1, Data Link_2
    "high_rssi": {"completed": None, "pending": [128]},
    "low_rssi": {"completed": None, "pending": [129]},
    "high_vswr": {"completed": None, "pending": [130]},
    "low_vswr": {"completed": None, "pending": [131]},
    "vswr_overthreshold": {"completed": None, "pending": [132]},

    # ---- Florida newly added cells ----
    "florida_header": 78,
    "florida_cells": list(range(79, 91)),

    # ---- Buffers ----
    "additional_completed": list(range(67, 77)),
    "additional_pending": list(range(133, 142)),
    "notes_final_port_config": 144,
    "notes_nr_verified": 145,
    "notes_buffer": list(range(146, 155)),
}
