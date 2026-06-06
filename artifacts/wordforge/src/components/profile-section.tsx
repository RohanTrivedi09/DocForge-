import { useState, useEffect } from "react";
import { ChevronDown, ChevronRight, User } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from "sonner";

export interface UserProfile {
  name: string;
  enrollNo: string;
  university: string;
  department: string;
  courseCode: string;
  subject: string;
}

const STORAGE_KEY = "docforge_profile";

export function loadProfile(): UserProfile | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function saveProfile(p: UserProfile) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(p));
}

const EMPTY: UserProfile = {
  name: "",
  enrollNo: "",
  university: "",
  department: "",
  courseCode: "",
  subject: "",
};

interface Props {
  onProfileLoad?: (profile: UserProfile) => void;
}

export function ProfileSection({ onProfileLoad }: Props) {
  const [open, setOpen] = useState(false);
  const [profile, setProfile] = useState<UserProfile>(EMPTY);

  useEffect(() => {
    const saved = loadProfile();
    if (saved) {
      setProfile(saved);
      onProfileLoad?.(saved);
    }
  }, []);

  const handleSave = () => {
    saveProfile(profile);
    onProfileLoad?.(profile);
    toast.success("Profile saved ✓", { duration: 2000 });
  };

  const set = (k: keyof UserProfile, v: string) =>
    setProfile((p) => ({ ...p, [k]: v }));

  return (
    <div className="rounded-lg border border-border bg-muted/20 overflow-hidden mb-2">
      <button
        className="w-full flex items-center gap-2 px-4 py-2.5 text-sm font-medium hover:bg-muted/40 transition-colors text-left"
        onClick={() => setOpen((o) => !o)}
      >
        <User className="w-3.5 h-3.5 text-muted-foreground" />
        <span className="flex-1 text-muted-foreground">My Profile</span>
        <span className="text-xs text-muted-foreground font-normal mr-1">auto-fills forms</span>
        {open ? (
          <ChevronDown className="w-3.5 h-3.5 text-muted-foreground" />
        ) : (
          <ChevronRight className="w-3.5 h-3.5 text-muted-foreground" />
        )}
      </button>

      {open && (
        <div className="px-4 pb-4 pt-1 space-y-3 border-t border-border bg-card">
          <div className="grid grid-cols-2 gap-3 pt-2">
            <div className="space-y-1">
              <Label className="text-xs">Full Name</Label>
              <Input className="h-8 text-sm" placeholder="Your name"
                value={profile.name} onChange={(e) => set("name", e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-xs">Enrollment No.</Label>
              <Input className="h-8 text-sm" placeholder="e.g. 1AUA23BCS159"
                value={profile.enrollNo} onChange={(e) => set("enrollNo", e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-xs">University</Label>
              <Input className="h-8 text-sm" placeholder="e.g. Adani University"
                value={profile.university} onChange={(e) => set("university", e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-xs">Department</Label>
              <Input className="h-8 text-sm" placeholder="e.g. Computer Science"
                value={profile.department} onChange={(e) => set("department", e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-xs">Course Code</Label>
              <Input className="h-8 text-sm" placeholder="e.g. ECSCI24305"
                value={profile.courseCode} onChange={(e) => set("courseCode", e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-xs">Subject</Label>
              <Input className="h-8 text-sm" placeholder="e.g. Deep Learning"
                value={profile.subject} onChange={(e) => set("subject", e.target.value)} />
            </div>
          </div>
          <Button size="sm" className="w-full h-8 text-xs" onClick={handleSave}>
            Save Profile
          </Button>
        </div>
      )}
    </div>
  );
}
