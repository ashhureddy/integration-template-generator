"""
Exact row map for Legacy_MCA (from direct inspection of the real template — row numbers
confirmed, not assumed). Each entry: completed_row(s), pending_row(s), stakeholder default.
Multi-slot items (Integration, Moved Sectors, Retune, FDD Renaming, Radio Swap, Transport SFP)
list all their available row slots — one used per detected instance, extras left unchecked.
"""

ROW_MAP = {
    # Subject / IWM / Configuration / IDL — single fixed rows, always the header value row (3)
    "subject": 3,
    "iwm_details": 6,
    "pre_configuration": 10,
    "current_configuration": 11,
    "post_configuration": 12,
    "wll_node": 13,
    "controller_6610": 14,
    "software_version": 15,
    "gs_version": 16,
    "idl_build_type": 19,
    "idle": 20,
    "idly": 21,
    "switch": 23,
    "slot_port": 24,

    # Completed section — multi-slot items list every available row
    "integration": {"completed": [42, 43, 44], "pending": [108, 109, 110], "stakeholder": "MIC PM"},
    "controller_integration": {"completed": [45], "pending": [111], "stakeholder": "MIC PM"},
    "port_conversion": {"completed": [46], "pending": [114], "stakeholder": "MIC PM"},
    "moved_sectors": {"completed": [47, 48, 49], "pending": None, "stakeholder": None},
    "deleted_node": {"completed": [50], "pending": None, "stakeholder": None},
    "deleted_sector": {"completed": [51], "pending": None, "stakeholder": None},
    "retune": {"completed": [52, 53], "pending": [115, 116], "stakeholder": "MIC"},
    "fdd_renaming": {"completed": [54, 55], "pending": [117, 118], "stakeholder": "MIC"},
    "radio_swap": {"completed": [56, 57, 58], "pending": [119, 120, 121], "stakeholder": "Tower Crew"},
    "dss_activation": {"completed": [59], "pending": [122], "stakeholder": "MIC|AT&T"},
    "ngs_activation": {"completed": [60], "pending": [123], "stakeholder": "MIC"},
    "gps_installation": {"completed": [61], "pending": [124], "stakeholder": "MIC PM|Tower Crew"},
    "lkf_installation": {"completed": [62], "pending": [125], "stakeholder": "MIC"},
    "psap_moved_lte": {"completed": [63], "pending": [126], "stakeholder": "MIC PM"},
    "speedtest_new_lte": {"completed": [64], "pending": [127], "stakeholder": "MIC PM"},
    "speedtest_5g": {"completed": [65], "pending": [128], "stakeholder": "MIC PM"},
    "calltest_fnet": {"completed": [66], "pending": [129], "stakeholder": "MIC PM"},
    "transport_sfp": {"completed": [67, 68, 69], "pending": [132, 133, 134, 135], "stakeholder": "MIC PM"},
    "sfp_installation": {"completed": [70], "pending": [130, 131], "stakeholder": "MIC PM/Tower Crew"},
    "ret_configuration": {"completed": [71], "pending": [136], "stakeholder": "Tower Crew"},
    "alarm_scripting": {"completed": [72], "pending": [137], "stakeholder": "MIC"},
    "sau_connections": {"completed": [73], "pending": [138], "stakeholder": "MIC PM"},
    "sup_connections": {"completed": [74], "pending": [139], "stakeholder": "MIC PM"},
    "xmu_installation": {"completed": [75], "pending": [140], "stakeholder": "MIC PM"},
    "idl_connections": {"completed": [76], "pending": [142], "stakeholder": "MIC PM"},
    "alarm_testing": {"completed": [77], "pending": [147], "stakeholder": "MIC PM"},
    "script_6673": {"completed": [78], "pending": [143], "stakeholder": "MIC PM"},
    "installation_generic": {"completed": [79], "pending": None, "stakeholder": None},
    "additional_completed": {"completed": list(range(81, 91)), "pending": None, "stakeholder": None},

    # Pending-only items (no Completed counterpart at all)
    "post_configuration_pending": {"completed": None, "pending": [107], "stakeholder": "MIC PM"},
    "siad_provisioning": {"completed": None, "pending": [112], "stakeholder": "AT&T"},
    "edp_publish": {"completed": None, "pending": [113], "stakeholder": "AT&T"},
    "rilinks_scripting": {"completed": None, "pending": [141], "stakeholder": "MIC PM"},
    "script_6673_config": {"completed": None, "pending": [144], "stakeholder": "AT&T"},
    "port_config_enm": {"completed": None, "pending": [145], "stakeholder": "AT&T"},
    "area_test": {"completed": None, "pending": [146], "stakeholder": "MIC PM"},
    "link_failure": {"completed": None, "pending": [148], "stakeholder": "Tower Crew"},
    "sfp_not_present": {"completed": None, "pending": [149], "stakeholder": "Tower Crew"},
    "mo_inconsistent_alarm": {"completed": None, "pending": [150], "stakeholder": "Tower Crew"},
    "fiberloss": {"completed": None, "pending": [151, 152], "stakeholder": "Tower Crew"},
    "high_rssi": {"completed": None, "pending": [153], "stakeholder": "Tower Crew"},
    "low_rssi": {"completed": None, "pending": [154], "stakeholder": "Tower Crew"},
    "high_vswr": {"completed": None, "pending": [155], "stakeholder": "Tower Crew"},
    "low_vswr": {"completed": None, "pending": [156], "stakeholder": "Tower Crew"},
    "vswr_overthreshold": {"completed": None, "pending": [157], "stakeholder": "Tower Crew"},
    "additional_pending": {"completed": None, "pending": list(range(158, 167)), "stakeholder": "varies"},

    # Pre-Existing Issues / Notes — generic repeated rows + Notes' 3 pre-canned templates
    "pre_existing_issues": list(range(169, 179)),
    "notes_final_port_config": 181,
    "notes_nr_verified": 182,
    "notes_mme_config": 183,   # has an xxNode_Idxx placeholder to substitute
    "notes_generic": list(range(184, 192)),
}
