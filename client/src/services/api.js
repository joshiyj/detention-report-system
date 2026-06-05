import axios from 'axios';

const API = axios.create({ baseURL: '/api' });

export async function generateReport(formData) {
  const { data } = await API.post('/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export async function generateReport2(formData) {
  const { data } = await API.post('/y2/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export async function generateReport3(formData) {
  const { data } = await API.post('/y3/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export async function generateReport4(formData) {
  const { data } = await API.post('/y4/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export async function generateReportBMED2(formData) {
  const { data } = await API.post('/bmed2/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export async function generateReportBMED3(formData) {
  const { data } = await API.post('/bmed3/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export function getDownloadUrl(filename) {
  return `/api/download/${filename}`;
}

