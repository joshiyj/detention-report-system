import axios from 'axios';

const API = axios.create({ baseURL: '/api' });

export async function generateReport(formData) {
  const { data } = await API.post('/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  return data;
}

export function getDownloadUrl(filename) {
  return `/api/download/${filename}`;
}
