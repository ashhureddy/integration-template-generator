"""
Row map for the real N2E_Pre_IX_Macro_V2 template (single sheet: N2E_Pre_IX).
Every row number below is taken directly from the actual confirmed dump of that file,
not assumed or carried over from the other (NSB-named) N2E file, which has different
row numbers entirely.
"""

N2E_ROW_MAP = {
    # ---- Header / Configuration / IDL Connections ----
    "subject": 3,                  # B=MIC(const) C=Market(const) D=IfN2E(const) E=Status(const) F=SiteName G=FACode H=SiteIDs
    "iwm_details": 6,
    "pre_configuration": 10,       # always "Nokia"
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
    "manual_feed_buffer": list(range(25, 40)),  # extra Sidehaul connections beyond the first

    # ---- Completed (41 = section header) ----
    "integration": {"completed": [42, 43, 44], "pending": None},
    "controller_integration": {"completed": [45], "pending": [79]},
    "dss_activation": {"completed": [46], "pending": [80]},
    "ngs_activation": {"completed": [47], "pending": [81]},
    "gps_installation": {"completed": [48], "pending": [82]},
    "lkf_installation": {"completed": [49], "pending": [83]},
    "transport_sfp": {"completed": [50, 51, 52], "pending": [87, 88, 89]},
    "ret_configuration": {"completed": [53], "pending": [90]},
    "external_alarm_scripting": {"completed": [54], "pending": [91]},
    "sau_connections": {"completed": [55], "pending": [92]},
    "sup_connections": {"completed": [56], "pending": [93]},
    "xmu_installation": {"completed": [57], "pending": [94]},
    "idl_connections": {"completed": [58], "pending": [95]},
    "smm_triggering": {"completed": [59], "pending": None},
    "area_test": {"completed": [60], "pending": [98]},
    "external_alarm_testing": {"completed": [61], "pending": [99]},
    "script_load_6673": {"completed": [62], "pending": None},
    "installation_manual": {"completed": [63], "pending": None},
    "sa_conversion": {"completed": [64], "pending": [104]},

    # ---- Pending-only, confirmed no Completed counterpart (per the real filled example) ----
    "psap_speedtest": {"completed": None, "pending": [84]},
    "speed_test_5g": {"completed": None, "pending": [85]},
    "calltest_fnet": {"completed": None, "pending": [86]},
    "on_site_nokia_cutover": {"completed": None, "pending": [78]},
    "config_6673": {"completed": None, "pending": [100]},
    "port_config_6673_enm": {"completed": None, "pending": [101]},
    "link_failure_or_sfp": {"completed": None, "pending": [102, 103]},
    "siad_provisioning": {"completed": None, "pending": [97]},
    "active_external_alarm": {"completed": None, "pending": None},

    # ---- Buffers ----
    "additional_completed": list(range(65, 74)),
    "additional_pending": list(range(105, 113)),
    "notes_final_port_config": 115,
    "notes_nr_verified": 116,
    "notes_buffer": list(range(117, 126)),
}
