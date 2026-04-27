SELECT WorkProjects.project_id, WorkProjects.project_name, WorkProjects.Description, WorkProjects.Municipality, WorkProjects.County, WorkProjects.cost, Val(Replace([project_id],"RTP-","")) AS SortID
FROM WorkProjects
WHERE WorkProjects.project_id LIKE "RTP-*"
ORDER BY WorkProjects.County, Val(Replace([project_id],"RTP-",""));

