SELECT InputProjects.[project_id], InputProjects.[project_name], InputProjects.[County], InputProjects.[Municipality], InputProjects.[Description], RTP_ExtendedData.County, RTP_ExtendedData.Municipality
FROM InputProjects LEFT JOIN RTP_ExtendedData ON InputProjects.project_id = RTP_ExtendedData.project_id;

