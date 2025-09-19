#!/usr/bin/env node

// Simple test to verify the MCP server can start without authentication
const { spawn } = require('child_process');
const path = require('path');

console.log('Testing Outlook MCP Server startup...');

// Set minimal environment for testing
const env = {
  ...process.env,
  OUTLOOK_CLIENT_ID: 'test-client-id',
  OUTLOOK_CLIENT_SECRET: 'test-client-secret',
  OUTLOOK_TENANT_ID: 'common'
};

const serverPath = path.join(__dirname, '../dist/index.js');
const child = spawn('node', [serverPath], {
  env,
  stdio: ['pipe', 'pipe', 'pipe']
});

let output = '';
let errorOutput = '';

child.stdout.on('data', (data) => {
  output += data.toString();
});

child.stderr.on('data', (data) => {
  errorOutput += data.toString();
});

// Send a simple test request
setTimeout(() => {
  const testRequest = {
    jsonrpc: '2.0',
    id: 1,
    method: 'tools/list',
    params: {}
  };
  
  child.stdin.write(JSON.stringify(testRequest) + '\n');
}, 1000);

// Kill after 3 seconds
setTimeout(() => {
  child.kill();
}, 3000);

child.on('close', (code) => {
  console.log('Server output:', errorOutput);
  console.log('Response output:', output);
  
  if (errorOutput.includes('Outlook MCP server running')) {
    console.log('✓ Server started successfully');
  } else {
    console.log('✗ Server startup may have issues');
  }
  
  if (output.includes('tools') || output.includes('jsonrpc')) {
    console.log('✓ Server responded to requests');
  } else {
    console.log('! Server response not detected (expected without proper auth)');
  }
  
  console.log('Test completed.');
});