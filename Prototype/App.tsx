import React, { useState, useEffect } from 'react';

const App: React.FC = () => {
  const [activeTab, setActiveTab] = useState('Lamp Rod');
  const [adminRole, setAdminRole] = useState(false);

  // Mock data for the product lines
  const productData = {
    'Lamp Rod': [
      { batchNumber: 'LR001', status: 'In Production', operator: 'Alice', timestamp: '2026-04-07 08:30:00' },
      { batchNumber: 'LR002', status: 'Completed', operator: 'Bob', timestamp: '2026-04-06 14:20:00' },
    ],
    'Baking Sand': [
      { batchNumber: 'BS001', status: 'In QA', operator: 'Catherine', timestamp: '2026-04-07 07:15:00' },
      { batchNumber: 'BS002', status: 'Scheduled', operator: 'David', timestamp: '2026-04-06 11:00:00' },
    ],
    'Serializing Materials': [
      { batchNumber: 'SM001', status: 'Pending', operator: 'Eve', timestamp: '2026-04-07 09:00:00' },
      { batchNumber: 'SM002', status: 'In Production', operator: 'Frank', timestamp: '2026-04-06 15:45:00' },
    ],
    'HV High Voltage Packing': [
      { batchNumber: 'HV001', status: 'Completed', operator: 'Grace', timestamp: '2026-04-06 10:30:00' },
      { batchNumber: 'HV002', status: 'In Production', operator: 'Hannah', timestamp: '2026-04-07 08:10:00' },
    ],
    Admin: [
      { username: 'alice', role: 'Admin' },
      { username: 'bob', role: 'User' },
    ]
  };

  const renderDataGrid = (tab: string) => {
    if (tab === 'Admin') {
      return adminRole ? (
        <table>
          <thead>
            <tr><th>Username</th><th>Role</th></tr></thead>
          <tbody>
            {productData.Admin.map((user, index) => (
              <tr key={index}><td>{user.username}</td><td>{user.role}</td></tr>
            ))}
          </tbody>
        </table>
      ) : <div>Access Denied</div>;
    }
    return (
      <table>
        <thead>
          <tr><th>Batch Number</th><th>Status</th><th>Operator</th><th>Timestamp</th></tr>
        </thead>
        <tbody>
          {productData[tab].map((product, index) => (
            <tr key={index}><td>{product.batchNumber}</td><td>{product.status}</td><td>{product.operator}</td><td>{product.timestamp}</td></tr>
          ))}
        </tbody>
      </table>
    );
  };

  return (
    <div style={{ fontFamily: 'Arial, sans-serif' }}>
      <nav style={{ marginBottom: '20px' }}>
        {['Lamp Rod', 'Baking Sand', 'Serializing Materials', 'HV High Voltage Packing', 'Admin'].map((tab) => (
          <button key={tab} onClick={() => setActiveTab(tab)} style={{
            backgroundColor: activeTab === tab ? '#007BFF' : '#fff',
            color: activeTab === tab ? '#fff' : '#000',
            border: '1px solid #007BFF',
            padding: '10px 20px',
            margin: '0 5px',
            cursor: 'pointer'
          }}>
            {tab}
          </button>
        ))}
      </nav>
      {renderDataGrid(activeTab)}
      {/* Connect backend here to fetch and update data */}
    </div>
  );
};

export default App;