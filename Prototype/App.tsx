// --- Professional Theme Style Constants ---
const proInput = {
  background: "#f8fafc", border: "1.3px solid #E5EAF0", borderRadius: 8,
  fontSize: 15, padding: "7px 12px", outline: "none", width: "100%",
  transition: "border .14s", boxShadow: "0 1px 2px #22336108",
};

const proTable = {
  width: "100%",
  background: "#fff",
  borderRadius: 13,
  boxShadow: "0 2px 10px #2233610f",
  border: "1px solid #E5EAF0",
  overflow: "hidden"
};

const proHeader = {
  background: "linear-gradient(90deg, #2563eb 60%, #10b981 140%)",
  color: "#fff", padding: "26px 0", borderBottomLeftRadius: 35, borderBottomRightRadius: 35,
  boxShadow: "0 4px 14px #1c223008",
  fontWeight: 900, fontSize: 28, textAlign: "center", letterSpacing: "-.01em"
};

const proPanel = {
  background: "#fff", borderRadius: 16, boxShadow: "0 2px 10px #2233610f", border: "1.5px solid #E5EAF0",
  padding: "30px 32px", marginBottom: 30
};

const proSectionTitle = {
  fontWeight: 800, fontSize: 22, marginBottom: 7, color: "#474e63", letterSpacing: "-.02em"
};

const proBtn = {
  background: "#2563eb", color: "#fff", fontSize: 16, padding: "10px 26px", fontWeight: 800,
  border: "none", borderRadius: 8, cursor: "pointer", boxShadow: "0 1.5px 7px #2563eb13", letterSpacing: ".01em"
};
// --- End Professional Theme Constants ---

import React from 'react';
import './App.css'; // Assuming global/professional CSS here, or remove if not needed

const App = () => {
    return (
        <div>
            <header style={proHeader}>Header</header>
            <main style={proPanel}>
                <section style={proSectionTitle}>Welcome to the App</section>
                <input type='text' style={proInput} placeholder='Enter text' />
                <table style={proTable}>
                    <tbody>
                        <tr style={{ background: '#f9f9f9' }}>
                            <td>Row 1</td>
                        </tr>
                        <tr>
                            <td>Row 2</td>
                        </tr>
                    </tbody>
                </table>
                <button style={proBtn}>Submit</button>
            </main>
        </div>
    );
};

export default App;