/* global Office, Excel, console, fetch */
import * as React from "react";
import { 
  Box, Button, Typography, Paper, Stack, 
  CircularProgress, Divider, Alert, Snackbar, 
  Tabs, Tab, TextField 
} from "@mui/material";
import AccountBalanceIcon from '@mui/icons-material/AccountBalance';
import TableChartIcon from '@mui/icons-material/TableChart';
const App: React.FC = () => {
  const [isOfficeReady, setIsOfficeReady] = React.useState<boolean>(false);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [platform, setPlatform] = React.useState<number>(Number(localStorage.getItem("active_platform")) || 0); 
  const [isConnectedQB, setIsConnectedQB] = React.useState<boolean>(!!localStorage.getItem("access_token_qb"));
  const [isConnectedXero, setIsConnectedXero] = React.useState<boolean>(!!localStorage.getItem("access_token_xero"));
  const [asOfDate, setAsOfDate] = React.useState(new Date().toISOString().split('T')[0]);
  const [toast, setToast] = React.useState({ open: false, message: "", severity: "success" as "success" | "error" });

  const primaryBlue = "#2c3e50";
  const accentAmber = "#f39221";

  const config = {
    qb: {
      id: "ABSW73Wyr1TLX9RJXaBOHGKOKQvG0oHWZHVDn2nyWbVB6Mp2qH",
      secret: "5dfIeIyGhC8vbbp84iXNdep1HSccoCLS51F2HtpU",
      scope: "com.intuit.quickbooks.accounting",
      auth: "https://appcenter.intuit.com/connect/oauth2"
    },
    xero: {
      id: "E84EC0577D2D422CBF638ED25B10910B",
      secret: "UrXJcfmxq0uSt5wBRNYt3n08FowTct2lLJGs7ZN1IzMaL-Va", 
      scope: "openid profile email accounting.reports.read offline_access",
      auth: "https://login.xero.com/identity/connect/authorize"
    }
  };

  const currentConfig = platform === 0 ? config.qb : config.xero;
  const isConnected = platform === 0 ? isConnectedQB : isConnectedXero;

  React.useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) setIsOfficeReady(true);
    });
  }, []);

  const handlePlatformChange = (_: any, newValue: number) => {
    setPlatform(newValue);
    localStorage.setItem("active_platform", newValue.toString());
  };

  const showToast = (msg: string, sev: "success" | "error") => setToast({ open: true, message: msg, severity: sev });

  const handleConnect = () => {
    setLoading(true);
    const redirectUri = window.location.origin + "/auth.html";
    const authUrl = `${currentConfig.auth}?client_id=${currentConfig.id}&response_type=code&scope=${encodeURIComponent(currentConfig.scope)}&redirect_uri=${redirectUri}&state=finbridge`;

    Office.context.ui.displayDialogAsync(authUrl, { height: 60, width: 40 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setLoading(false);
        showToast("Login failed.", "error");
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
        try {
          const msg = JSON.parse(arg.message);
          if (msg.status === "success") {
            if (platform === 0) localStorage.setItem("qb_realmId", msg.realmId);
            exchangeToken(msg.code);
          }
        } catch (e) { console.error(e); }
        dialog.close();
      });
    });
  };

  const exchangeToken = async (code: string) => {
    const proxyUrl = "https://cors-anywhere.herokuapp.com/";
    const tokenUrl = platform === 0 ? "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer" : "https://identity.xero.com/connect/token";
    try {
      const response = await fetch(proxyUrl + tokenUrl, {
        method: "POST",
        headers: { "Authorization": `Basic ${btoa(currentConfig.id + ":" + currentConfig.secret)}`, "Content-Type": "application/x-www-form-urlencoded", "X-Requested-With": "XMLHttpRequest" },
        body: new URLSearchParams({ grant_type: "authorization_code", code: code, redirect_uri: window.location.origin + "/auth.html" })
      });
      const data = await response.json();
      if (data.access_token) {
        const tokenKey = platform === 0 ? "access_token_qb" : "access_token_xero";
        localStorage.setItem(tokenKey, data.access_token);
        if (platform === 1) await fetchXeroTenantId(data.access_token);
        platform === 0 ? setIsConnectedQB(true) : setIsConnectedXero(true);
        showToast("Connected!", "success");
      }
    } catch (e) { showToast("Auth failed.", "error"); }
    finally { setLoading(false); }
  };

  const fetchXeroTenantId = async (token: string) => {
    try {
      const res = await fetch("https://cors-anywhere.herokuapp.com/https://api.xero.com/connections", {
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" }
      });
      const connections = await res.json();
      if (connections.length > 0) localStorage.setItem("xero_tenant_id", connections[0].tenantId);
    } catch (e) { console.error(e); }
  };

  const syncTrialBalance = async () => {
    setLoading(true);
    const token = localStorage.getItem(platform === 0 ? "access_token_qb" : "access_token_xero");
    const proxyUrl = "https://cors-anywhere.herokuapp.com/";
    let reportUrl = "";
    let headers: any = { "Authorization": `Bearer ${token}`, "Accept": "application/json", "X-Requested-With": "XMLHttpRequest" };

    if (platform === 0) {
      const rId = localStorage.getItem("qb_realmId");
      reportUrl = `https://sandbox-quickbooks.api.intuit.com/v3/company/${rId}/reports/TrialBalance?end_date=${asOfDate}`;
    } else {
      const tId = localStorage.getItem("xero_tenant_id");
      reportUrl = `https://api.xero.com/api.xro/2.0/Reports/TrialBalance?date=${asOfDate}`;
      headers["xero-tenant-id"] = tId;
    }

    try {
      const response = await fetch(proxyUrl + reportUrl, { method: "GET", headers: headers });
      if (!response.ok) throw new Error("API failed");
      const data = await response.json();
      await writeToExcelMaster(data);
      showToast("Data Imported Successfully!", "success");
    } catch (error) {
      console.error(error);
      showToast("Sync Error. Check Proxy Access.", "error");
    } finally { setLoading(false); }
  };

  const writeToExcelMaster = async (data: any) => {
    try {
      await Excel.run(async (context) => {
        const sheetName = platform === 0 ? "TB_DATA_QB" : "TB_DATA_XERO";
        let sheets = context.workbook.worksheets;
        let sheet = sheets.getItemOrNullObject(sheetName);
        await context.sync();

        if (sheet.isNullObject) {
          sheet = sheets.add(sheetName);
          await context.sync();
        }
        
        sheet.activate();
        sheet.getRange("A1:C500").clear();

        let tableData: any[][] = [["Account", "Debit", "Credit"]];

        if (platform === 0) {
          const rows = data.Rows?.Row || [];
          rows.forEach((r: any) => {
            if (r.ColData) tableData.push([r.ColData[0].value, r.ColData[1].value || "0", r.ColData[2].value || "0"]);
          });
        } else {
          // --- XERO DEEP PARSING (Recursive) ---
          const extractXeroRows = (rows: any[]) => {
            rows.forEach(row => {
              if (row.Cells && row.Cells.length >= 3 && row.RowType === "Row") {
                tableData.push([row.Cells[0].Value, row.Cells[1].Value || "0", row.Cells[2].Value || "0"]);
              }
              if (row.Rows) extractXeroRows(row.Rows); // Check for nested rows
            });
          };
          if (data.Reports && data.Reports[0].Rows) extractXeroRows(data.Reports[0].Rows);
        }

        if (tableData.length > 1) {
          const range = sheet.getRangeByIndexes(0, 0, tableData.length, 3);
          range.values = tableData;
          range.getRow(0).format.fill.color = primaryBlue;
          range.getRow(0).format.font.color = "white";
          range.getRow(0).format.font.bold = true;
          sheet.getRange("A:C").format.autofitColumns();
        }
        await context.sync();
      });
    } catch (e) {
      console.error("Excel Error:", e);
    }
  };

  return (
    <Box sx={{ bgcolor: "#ffffff", minHeight: "100vh" }}>
      <Box sx={{ p: 2, textAlign: "center", borderBottom: `2px solid ${primaryBlue}` }}>
        <AccountBalanceIcon sx={{ fontSize: 35, color: primaryBlue }} />
        <Typography sx={{ fontSize: "14px", fontWeight: 900, color: primaryBlue }}>FINBRIDGE PRO</Typography>
      </Box>

      <Tabs value={platform} onChange={handlePlatformChange} centered sx={{ bgcolor: "#f8f9fa" }}>
        <Tab label="QuickBooks" sx={{ fontSize: '12px', fontWeight: 'bold' }} />
        <Tab label="Xero" sx={{ fontSize: '12px', fontWeight: 'bold' }} />
      </Tabs>

      <Box sx={{ p: 2.5 }}>
        <Stack spacing={2}>
          <Paper variant="outlined" sx={{ p: 1.5, bgcolor: "#fcfcfc" }}>
            <Typography sx={{ fontSize: "8.5px", fontWeight: 800, color: "#999", mb: 1, textTransform: 'uppercase' }}>Parameters</Typography>
            <TextField type="date" label="Report Date" fullWidth size="small" value={asOfDate} onChange={(e) => setAsOfDate(e.target.value)} InputLabelProps={{ shrink: true }} />
          </Paper>

          {!isConnected ? (
            <Button fullWidth variant="contained" onClick={handleConnect} disabled={loading} sx={{ bgcolor: primaryBlue, py: 1.2, fontWeight: 700, textTransform: 'none' }}>
              {loading ? <CircularProgress size={20} color="inherit" /> : `Connect ${platform === 0 ? "QuickBooks" : "Xero"}`}
            </Button>
          ) : (
            <Stack spacing={1}>
              <Button fullWidth variant="contained" onClick={syncTrialBalance} disabled={loading} startIcon={loading ? <CircularProgress size={18} color="inherit" /> : <TableChartIcon />} sx={{ bgcolor: "#2ecc71", py: 1.2, fontWeight: 700, textTransform: 'none' }}>
                Pull Trial Balance
              </Button>
              <Button variant="text" size="small" sx={{ fontSize: '9px' }} onClick={() => { localStorage.clear(); window.location.reload(); }}>DISCONNECT</Button>
            </Stack>
          )}
        </Stack>
        
        <Box sx={{ mt: 4, p: 1.5, bgcolor: "#333", borderRadius: 1 }}>
           <Typography sx={{ fontSize: '9px', color: accentAmber, fontWeight: 800 }}>ENGINE STATUS</Typography>
           <Typography sx={{ fontSize: '10px', color: '#90caf9', fontFamily: 'monospace', mt: 1 }}>
             {isConnected ? "> Secure link active." : "> Awaiting authorization..."}
             {loading && <><br/> Processing...</>}
           </Typography>
        </Box>
      </Box>

      <Snackbar open={toast.open} autoHideDuration={4000} onClose={() => setToast({ ...toast, open: false })} anchorOrigin={{ vertical: "bottom", horizontal: "center" }}>
        <Alert severity={toast.severity} variant="filled" sx={{ width: '100%', fontSize: '11px' }}>{toast.message}</Alert>
      </Snackbar>
    </Box>
  );
};

export default App;