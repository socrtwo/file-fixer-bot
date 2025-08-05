-- Create storage bucket for file uploads
INSERT INTO storage.buckets (id, name, public) VALUES ('file-repairs', 'file-repairs', false);

-- Create storage policies for file access
CREATE POLICY "Allow file uploads" ON storage.objects
FOR INSERT WITH CHECK (bucket_id = 'file-repairs');

CREATE POLICY "Allow file downloads" ON storage.objects
FOR SELECT USING (bucket_id = 'file-repairs');

CREATE POLICY "Allow file updates" ON storage.objects
FOR UPDATE USING (bucket_id = 'file-repairs');

CREATE POLICY "Allow file deletions" ON storage.objects
FOR DELETE USING (bucket_id = 'file-repairs');