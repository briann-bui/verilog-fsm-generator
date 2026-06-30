module {{module_name}} (
{{port_declarations}}
);

{{state_localparams}}

reg [STATE_W-1:0] {{state_reg_name}};
reg [STATE_W-1:0] {{next_state_name}};

{{state_register_block}}

{{comb_block}}

endmodule
