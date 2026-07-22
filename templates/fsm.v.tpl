module {{module_name}} (
{{port_declarations}}
);

{{state_localparams}}

reg [STATE_W-1:0] {{state_reg_name}};
reg [STATE_W-1:0] {{next_state_name}};

{{timer_declarations}}

{{state_register_block}}

{{timer_block}}

{{comb_block}}

endmodule
