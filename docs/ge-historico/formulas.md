# \# Fórmulas GE - Histórico

# 

# \## A1

# A1 deve ficar igual a D63.

# 

# \## C1

# Usar esta fórmula localizada:

# 

# =MÉDIA(D3:D23)\*MÉDIA(SEERRO(FILTER(H63:LH63;MOD(COL(H63:LH63)-COL(H63);4)=0);""))/MÉDIA(SEERRO(FILTER(H3:LH23;MOD(COL(H3:LH23)-COL(H3);4)=0);""))

