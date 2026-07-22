`default_nettype none

module {{module_name}} (
{{port_declarations}}
);

{{state_declarations}}

{{timer_declarations}}

{{state_register_block}}

{{timer_block}}

{{comb_block}}

{{assertion_block}}

endmodule

`default_nettype wire
