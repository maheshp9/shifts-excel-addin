import axios from 'axios';

export const getGraphData = async (url: string, accesstoken: string) => {
    const response = await axios({
        url: url,
        method: 'get',
        headers: {'Authorization': `Bearer ${accesstoken}`}
      });
    return response;
    // TODO: Handle nextlink and return the data, instead of Axiosresponse
};

export const updateGraphData = async (url: string, data: any, accesstoken: string) => {
  return await axios({
    url: url,
    method: 'put',
    headers: { 'Authorization': `Bearer ${accesstoken}`, 'Content-Type': 'application/json'},
    data: data
  });
};

export const postGraphData = async (url: string, data: any, accesstoken: string) => {
  return await axios({
    url: url,
    method: 'post',
    headers: { 'Authorization': `Bearer ${accesstoken}`, 'Content-Type': 'application/json'},
    data: data
  });
};

export const deleteGraphData = async (url: string, data: any, accesstoken: string) => {
  return await axios({
    url: url,
    method: 'delete',
    headers: { 'Authorization': `Bearer ${accesstoken}`, 'Content-Type': 'application/json'},
    data: data
  });
};
